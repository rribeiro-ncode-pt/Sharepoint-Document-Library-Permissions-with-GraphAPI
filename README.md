# SharePoint Permissions Exporter

A .NET 8 console application that retrieves and exports file-level permissions from SharePoint Document Libraries to CSV format using the Microsoft Graph API.

## Overview

This application provides a comprehensive solution for auditing and exporting SharePoint document library permissions. It recursively traverses document libraries, retrieves detailed permission information for each file, and exports the data to a CSV file for analysis.

### Key Features

- **Azure AD Authentication**: Secure authentication using Azure AD App Registration with client credentials flow
- **Recursive Traversal**: Automatically processes all files in a document library, including nested folders
- **Batch Processing**: Efficiently handles large libraries using Microsoft Graph batch requests (20 requests per batch)
- **Retry Logic**: Automatic retry with exponential backoff for handling API throttling and transient errors
- **Progress Tracking**: Real-time console output showing processing progress
- **Detailed Export**: Exports comprehensive permission data including roles, users/groups, and inheritance information
- **Configurable**: Flexible configuration options for batch size, delays, and retry behavior

## Prerequisites

Before running this application, ensure you have the following:

### Required Software

- **.NET 8 SDK**: Download and install from [https://dotnet.microsoft.com/download/dotnet/8.0](https://dotnet.microsoft.com/download/dotnet/8.0)

### Azure AD App Registration

You need an Azure AD App Registration with the following:

- **Required API Permissions** (Microsoft Graph):
  - `Sites.Read.All` - Application permission
  - `Files.Read.All` - Application permission
  
- **Admin Consent**: These application permissions require admin consent to be granted by a tenant administrator

- **Client Secret**: A valid client secret must be created for the app registration

## Azure AD App Registration Setup

Follow these steps to create and configure an Azure AD App Registration:

### 1. Create App Registration

1. Sign in to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Enter a name (e.g., "SharePoint Permissions Exporter")
5. Select **Accounts in this organizational directory only**
6. Click **Register**

### 2. Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph** > **Application permissions**
4. Search for and add:
   - `Sites.Read.All`
   - `Files.Read.All`
5. Click **Add permissions**
6. Click **Grant admin consent for [Your Organization]** (requires admin rights)
7. Confirm the consent

### 3. Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Enter a description (e.g., "SharePoint Exporter Secret")
4. Select an expiration period
5. Click **Add**
6. **Important**: Copy the secret **Value** immediately - it won't be shown again

### 4. Gather Required Information

You'll need three values for configuration:

- **Tenant ID**: Found in **Azure Active Directory** > **Overview** > **Tenant ID**
- **Client ID** (Application ID): Found in your app registration's **Overview** page
- **Client Secret**: The value you copied in step 3

## Configuration

The application uses the [`appsettings.json`](appsettings.json:1) file for configuration. Update this file with your specific values.

### Configuration File Structure

```json
{
  "AzureAd": {
    "TenantId": "your-tenant-id-here",
    "ClientId": "your-client-id-here",
    "ClientSecret": "your-client-secret-here"
  },
  "SharePoint": {
    "SiteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
    "DocumentLibraryName": ""
  },
  "Export": {
    "OutputFileName": "SharePointPermissions.csv",
    "DelayBetweenRequestsMs": 100,
    "BatchSize": 20,
    "MaxRetryAttempts": 3
  }
}
```

### Configuration Settings

#### AzureAd Section

| Setting | Required | Description |
|---------|----------|-------------|
| `TenantId` | Yes | Your Azure AD tenant ID (GUID format) |
| `ClientId` | Yes | Application (client) ID from your app registration |
| `ClientSecret` | Yes | Client secret value created in Azure AD |

#### SharePoint Section

| Setting | Required | Description |
|---------|----------|-------------|
| `SiteUrl` | Yes | Full URL to your SharePoint site (e.g., `https://contoso.sharepoint.com/sites/Finance`) |
| `DocumentLibraryName` | No | Name of the document library to process. Leave empty or omit to use the default document library |

#### Export Section

| Setting | Required | Default | Description |
|---------|----------|---------|-------------|
| `OutputFileName` | No | `SharePointPermissions.csv` | Name of the output CSV file |
| `DelayBetweenRequestsMs` | No | `100` | Delay in milliseconds between API requests (helps avoid throttling) |
| `BatchSize` | No | `20` | Number of requests per batch (Graph API supports up to 20) |
| `MaxRetryAttempts` | No | `3` | Maximum number of retry attempts for failed requests |

### Configuration Example

```json
{
  "AzureAd": {
    "TenantId": "12345678-1234-1234-1234-123456789012",
    "ClientId": "87654321-4321-4321-4321-210987654321",
    "ClientSecret": "abc123XYZ~def456-ghi789.jkl012"
  },
  "SharePoint": {
    "SiteUrl": "https://contoso.sharepoint.com/sites/Finance",
    "DocumentLibraryName": "Shared Documents"
  },
  "Export": {
    "OutputFileName": "FinancePermissions.csv",
    "DelayBetweenRequestsMs": 150,
    "BatchSize": 20,
    "MaxRetryAttempts": 5
  }
}
```

## Installation & Setup

### 1. Download/Clone the Repository

Download or clone this repository to your local machine:

```bash
git clone <repository-url>
cd SharePointPermissionsExporter
```

### 2. Restore Dependencies

Restore the required NuGet packages:

```bash
dotnet restore
```

### 3. Configure Application

Edit [`appsettings.json`](appsettings.json:1) with your Azure AD and SharePoint details (see [Configuration](#configuration) section).

### 4. Build the Project

Build the application to ensure everything is configured correctly:

```bash
dotnet build
```

If the build is successful, you'll see:

```
Build succeeded.
    0 Warning(s)
    0 Error(s)
```

## Usage

### Running the Application

Execute the application from the project directory:

```bash
dotnet run
```

### Execution Flow

The application will:

1. **Load Configuration**: Validates all required settings from [`appsettings.json`](appsettings.json:1)
2. **Authenticate**: Establishes a connection to Microsoft Graph API using Azure AD credentials
3. **Retrieve Site Information**: Gets the SharePoint site ID and document library (drive) ID
4. **Fetch Permissions**: Recursively processes all files and retrieves permission information
5. **Export to CSV**: Writes all permission data to the specified output file
6. **Display Summary**: Shows statistics about the export

### Sample Output

```
=================================================
SharePoint Document Library Permissions Exporter
=================================================

Configuration loaded successfully:
  - Site URL: https://contoso.sharepoint.com/sites/Finance
  - Document Library: Shared Documents
  - Output File: SharePointPermissions.csv
  - Batch Size: 20
  - Delay Between Requests: 100ms
  - Max Retry Attempts: 3

Step 1: Authenticating with Microsoft Graph API...
✓ Authentication successful

Step 2: Retrieving SharePoint site and drive information...
✓ Site and drive information retrieved

Step 3: Fetching file permissions (this may take a while)...
Processing files: 245/245 (100.0%)
✓ Retrieved 1,234 permission records

Step 4: Exporting permissions to CSV...
✓ Successfully exported 1,234 records to SharePointPermissions.csv

=================================================
Summary Statistics
=================================================
Total Permissions Exported: 1,234
Unique Files: 245
Inherited Permissions: 890
Direct Permissions: 344
Unique Users/Groups: 42
Total Execution Time: 45.23 seconds
=================================================

Top 5 Most Common Roles:
  - read: 756 occurrences
  - write: 389 occurrences
  - owner: 89 occurrences
```

### Output Location

The CSV file will be created in the same directory as the application executable:
- When running with `dotnet run`: Project root directory
- When running compiled executable: Same directory as the `.exe` file

## Output Format

The exported CSV file contains the following columns:

| Column Name | Description | Example |
|-------------|-------------|---------|
| `FileName` | Name of the file | `Q4-Report.xlsx` |
| `WebUrl` | Full URL to access the file in SharePoint | `https://contoso.sharepoint.com/sites/Finance/...` |
| `FileId` | Unique identifier for the file | `01ABCDEF123456789` |
| `PermissionId` | Unique identifier for this specific permission | `aTowIy5mfG1...` |
| `RolesString` | Comma-separated list of permission roles | `read, write` |
| `GrantedToDisplayName` | Display name of the user or group | `John Doe` or `Finance Team` |
| `GrantedToEmail` | Email address (if available) | `john.doe@contoso.com` |
| `IsInherited` | Whether permission is inherited from parent | `True` or `False` |
| `InheritedFrom` | Path/URL from which permission is inherited | `/sites/Finance/Shared Documents` |

### Understanding the Data

- **Multiple rows per file**: Each file may have multiple rows if it has permissions granted to multiple users/groups
- **Inherited permissions**: When `IsInherited` is `True`, the permission comes from a parent folder. The `InheritedFrom` column shows the source
- **Direct permissions**: When `IsInherited` is `False`, the permission is set directly on this file
- **Roles**: Common roles include:
  - `read` - View only access
  - `write` - Edit access
  - `owner` - Full control
  - `sp.full control` - SharePoint-specific full control
  - `sp.limited access` - Limited access (common for inherited permissions)

### Sample CSV Output

```csv
FileName,WebUrl,FileId,PermissionId,RolesString,GrantedToDisplayName,GrantedToEmail,IsInherited,InheritedFrom
Budget-2024.xlsx,https://contoso.sharepoint.com/...,01ABC123,perm1,"read, write",Finance Team,finance@contoso.com,False,
Budget-2024.xlsx,https://contoso.sharepoint.com/...,01ABC123,perm2,read,John Doe,john.doe@contoso.com,True,/sites/Finance/Shared Documents
Report.docx,https://contoso.sharepoint.com/...,01XYZ789,perm3,owner,Jane Smith,jane.smith@contoso.com,False,
```

## Troubleshooting

### Authentication Errors

**Problem**: `❌ Invalid credentials` or authentication failures

**Solutions**:
- Verify your `TenantId`, `ClientId`, and `ClientSecret` are correct
- Ensure the client secret hasn't expired (check in Azure AD)
- Confirm admin consent has been granted for the API permissions
- Check that the app registration is active (not disabled)

### Missing Permissions Error

**Problem**: Error messages about insufficient permissions or access denied

**Solutions**:
- Verify `Sites.Read.All` and `Files.Read.All` permissions are added as **Application permissions** (not Delegated)
- Ensure **admin consent** has been granted (look for green checkmarks in Azure AD)
- Wait a few minutes after granting consent for changes to propagate

### Throttling (429 Errors)

**Problem**: Application encounters HTTP 429 (Too Many Requests) errors

**Solutions**:
- The application automatically retries with exponential backoff
- Increase `DelayBetweenRequestsMs` in [`appsettings.json`](appsettings.json:1) (e.g., to 200 or 300)
- Decrease `BatchSize` to process fewer items per request
- Increase `MaxRetryAttempts` for more persistent retry behavior

### Empty Results

**Problem**: `⚠️ No permissions found` message

**Solutions**:
- Verify the `SiteUrl` is correct and accessible
- Check the `DocumentLibraryName` if specified (try leaving it empty for default library)
- Ensure the document library actually contains files
- Confirm the Azure AD app has proper permissions and consent
- Try accessing the SharePoint site manually to verify it exists

### Build Errors

**Problem**: Build fails or dependencies not found

**Solutions**:
- Ensure .NET 8 SDK is installed: `dotnet --version` (should show 8.x.x)
- Run `dotnet restore` to restore NuGet packages
- Delete `bin` and `obj` folders, then rebuild
- Check for internet connectivity (required for NuGet package restoration)

### Configuration Validation Failed

**Problem**: `❌ Configuration validation failed`

**Solutions**:
- Ensure all placeholder values in [`appsettings.json`](appsettings.json:1) are replaced
- Remove any text containing `"your-"` or `"-here"`
- Verify JSON syntax is valid (use a JSON validator if needed)
- Check that required fields are not empty or whitespace

## Technical Details

### Architecture Overview

The application is structured with a clean separation of concerns:

- **[`Program.cs`](Program.cs:1)**: Main entry point, orchestrates the workflow and handles configuration
- **[`AuthenticationService.cs`](Services/AuthenticationService.cs:1)**: Manages Azure AD authentication and Graph client creation
- **[`SharePointPermissionsService.cs`](Services/SharePointPermissionsService.cs:1)**: Handles all SharePoint and Microsoft Graph API interactions
- **[`CsvExportService.cs`](Services/CsvExportService.cs:1)**: Manages CSV file generation using CsvHelper library
- **[`FilePermissionInfo.cs`](Models/FilePermissionInfo.cs:1)**: Data model representing file permission information

### Batch Processing

The application uses Microsoft Graph batch requests to optimize API performance:

- Groups up to **20 permission requests** into a single batch request
- Significantly reduces total API calls and execution time
- Batch size is configurable via `BatchSize` setting
- Each batch request counts as 1 API call + number of individual requests in the batch

### Retry Logic

Implements exponential backoff for handling transient errors:

1. **Initial attempt**: Try the API request
2. **On failure**: Wait for an exponentially increasing delay
3. **Retry**: Attempt the request again
4. **Repeat**: Continue until success or max retry attempts reached

Delay calculation: `delay = baseDelay * (2 ^ attemptNumber)`

This approach is particularly effective for handling:
- HTTP 429 (Too Many Requests) throttling errors
- Temporary network issues
- Transient service unavailability

### Performance Considerations

For large document libraries (1000+ files):

- **Execution Time**: Expect 1-2 seconds per batch of 20 files
- **Memory Usage**: The application loads all permissions into memory before export
- **API Throttling**: Microsoft Graph has usage limits; the built-in delay and retry logic helps stay within limits
- **Recommendations**:
  - Process libraries during off-peak hours
  - Increase `DelayBetweenRequestsMs` for very large libraries
  - Monitor progress output to track processing speed

### Dependencies

The application uses the following NuGet packages:

- **Microsoft.Graph** (v5.86.0) - Microsoft Graph SDK for API access
- **Azure.Identity** (v1.13.1) - Azure authentication library
- **CsvHelper** (v33.0.1) - CSV reading and writing library
- **Microsoft.Extensions.Configuration** (v8.0.0) - Configuration management
- **Microsoft.Extensions.Configuration.Json** (v8.0.1) - JSON configuration support

## License

This project is provided as-is for internal use. Modify and distribute according to your organization's policies.

---

**Note**: This application requires appropriate permissions and access to SharePoint. Always ensure you have proper authorization before running permission audits on organizational data.