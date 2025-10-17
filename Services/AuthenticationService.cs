using Azure.Identity;
using Microsoft.Graph;

namespace SharePointPermissionsExporter.Services;

/// <summary>
/// Handles authentication with Microsoft Graph API using client credentials
/// </summary>
public class AuthenticationService
{
    private readonly string _tenantId;
    private readonly string _clientId;
    private readonly string _clientSecret;

    /// <summary>
    /// Initializes a new instance of the AuthenticationService
    /// </summary>
    /// <param name="tenantId">Azure AD Tenant ID</param>
    /// <param name="clientId">Azure AD Application (Client) ID</param>
    /// <param name="clientSecret">Azure AD Application Client Secret</param>
    public AuthenticationService(string tenantId, string clientId, string clientSecret)
    {
        _tenantId = tenantId ?? throw new ArgumentNullException(nameof(tenantId));
        _clientId = clientId ?? throw new ArgumentNullException(nameof(clientId));
        _clientSecret = clientSecret ?? throw new ArgumentNullException(nameof(clientSecret));
    }

    /// <summary>
    /// Creates and returns an authenticated GraphServiceClient using client credentials flow
    /// </summary>
    /// <returns>Authenticated GraphServiceClient instance</returns>
    public GraphServiceClient GetAuthenticatedGraphClient()
    {
        try
        {
            // Create client secret credential
            var clientSecretCredential = new ClientSecretCredential(
                _tenantId,
                _clientId,
                _clientSecret
            );

            // Define the required scopes for Microsoft Graph
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // Create and return the GraphServiceClient
            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            Console.WriteLine("Successfully authenticated with Microsoft Graph API");
            return graphClient;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Authentication failed: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Validates that the authentication credentials are not empty
    /// </summary>
    /// <returns>True if credentials are valid, false otherwise</returns>
    public bool ValidateCredentials()
    {
        if (string.IsNullOrWhiteSpace(_tenantId))
        {
            Console.WriteLine("Error: TenantId is missing or empty");
            return false;
        }

        if (string.IsNullOrWhiteSpace(_clientId))
        {
            Console.WriteLine("Error: ClientId is missing or empty");
            return false;
        }

        if (string.IsNullOrWhiteSpace(_clientSecret))
        {
            Console.WriteLine("Error: ClientSecret is missing or empty");
            return false;
        }

        return true;
    }
}