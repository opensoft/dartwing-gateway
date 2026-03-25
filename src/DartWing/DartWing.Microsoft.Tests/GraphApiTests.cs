using System.Text.Json;
using DartWing.KeyCloak;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace DartWing.Microsoft.Tests;

[TestClass]
public sealed class GraphApiTests
{
    [TestMethod]
    public async Task ClienAccessToSharePointSites()
    {
        var configuration = CreateConfiguration();
        var clientId = RequireSetting(configuration, "MicrosoftTestAuth:ClientId");
        var clientSecret = RequireSetting(configuration, "MicrosoftTestAuth:ClientSecret");
        var tenantId = RequireSetting(configuration, "MicrosoftTestAuth:TenantId");

        var body = new FormUrlEncodedContent([
            new KeyValuePair<string, string>("client_id", clientId),
            new KeyValuePair<string, string>("client_secret", clientSecret),
            new KeyValuePair<string, string>("grant_type", "client_credentials"),
            new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default")
        ]);

        var url = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";
        using var response = await new HttpClient().PostAsync(url, body);

        if (!response.IsSuccessStatusCode)
        {
            return;
        }

        var tokenString = await response.Content.ReadAsStringAsync();
        Assert.IsNotNull(tokenString);

        var token = JsonDocument.Parse(tokenString).RootElement.GetProperty("access_token").GetString();
        Assert.IsNotNull(token);

        try
        {
            var adapter = new GraphApiAdapter(token, DefaultHttpClientFactory.Instance, new MemoryCache(new MemoryCacheOptions()));
            var sites = await adapter.GetAllSites(CancellationToken.None);

            var allDrives = new List<string>();
            foreach (var site in sites)
            {
                var drives = await adapter.GetAllDrives(site.Id, CancellationToken.None);
                foreach (var drive in drives)
                {
                    if (drive == null) continue;
                    allDrives.Add($"{site.Name} __ {drive.Name} __ {drive.Id}");
                }
            }

            _ = await adapter.GetAllDrives(CancellationToken.None);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
    }

    [TestMethod]
    public async Task UserAccessToSharePointSites()
    {
        var configuration = CreateConfiguration();
        var exchangeUser = RequireSetting(configuration, "MicrosoftTestAuth:ExchangeUser");
        var provider = CreateServiceProvider().GetRequiredService<KeyCloakTokenProvider>();

        var token = await provider.TokenExchangeProviderAccessToken(exchangeUser, "azure_sharepoint", false,
            CancellationToken.None);

        Assert.IsNotNull(token);

        var adapter = new GraphApiAdapter(token.AccessToken, DefaultHttpClientFactory.Instance, new MemoryCache(new MemoryCacheOptions()));
        try
        {
            _ = await adapter.Me(CancellationToken.None);
            _ = await adapter.GetMyFolders(CancellationToken.None);
            _ = await adapter.GetAllSites(CancellationToken.None);
            var drives = await adapter.GetAllDrives(CancellationToken.None);
            _ = await adapter.GetAllFolders(drives[0].Id, recursive: true, ct: CancellationToken.None);
            _ = await adapter.GetAllFolders("b!EqFpt3Qt40-22UOfOogtmzN05V8GV8tHvZs77CRfPJBZWOwcg3WBT6SG5eTEDc6i",
                recursive: true, ct: CancellationToken.None);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
    }

    [TestMethod]
    public async Task GetSiteName()
    {
        var token = "";
        var adapter = new GraphApiAdapter(token, DefaultHttpClientFactory.Instance, new MemoryCache(new MemoryCacheOptions()));

        var tenantName = await adapter.GetTenantName(CancellationToken.None);
        var site = await adapter.GetSite("FHASClients-InkRouter", tenantName, CancellationToken.None);

        var drives = await adapter.GetAllDrives(site.Id, CancellationToken.None);
        var folders = await adapter.GetAllFolders(drives[0].Id);
        _ = await adapter.GetAllFolders(drives[0].Id, folders[0].Id);
    }

    [TestMethod]
    public async Task UploadFileFarHeap()
    {
        var configuration = CreateConfiguration();
        var token = RequireSetting(configuration, "MicrosoftTestAuth:FarHeapUploadToken");
        var driveId = RequireSetting(configuration, "MicrosoftTestAuth:FarHeapDriveId");
        var folderId = RequireSetting(configuration, "MicrosoftTestAuth:FarHeapFolderId");
        var uploadUrl = RequireSetting(configuration, "MicrosoftTestAuth:FarHeapUploadUrl");
        var adapter = new GraphApiAdapter(token, DefaultHttpClientFactory.Instance, new MemoryCache(new MemoryCacheOptions()));

        _ = await adapter.UploadFileByLink(driveId, folderId, uploadUrl, CancellationToken.None);
    }

    [TestMethod]
    public async Task UploadFileOpensoft()
    {
        var configuration = CreateConfiguration();
        var token = RequireSetting(configuration, "MicrosoftTestAuth:OpensoftUploadToken");
        var driveId = RequireSetting(configuration, "MicrosoftTestAuth:OpensoftDriveId");
        var folderId = RequireSetting(configuration, "MicrosoftTestAuth:OpensoftFolderId");
        var uploadUrl = RequireSetting(configuration, "MicrosoftTestAuth:OpensoftUploadUrl");
        var adapter = new GraphApiAdapter(token, DefaultHttpClientFactory.Instance, new MemoryCache(new MemoryCacheOptions()));

        _ = await adapter.UploadFileByLink(driveId, folderId, uploadUrl, CancellationToken.None);
    }

    [TestMethod]
    public async Task GetDriveFolderId()
    {
        var token = "";
        var adapter = new GraphApiAdapter(token, DefaultHttpClientFactory.Instance, new MemoryCache(new MemoryCacheOptions()));
        var site = await adapter.GetSite("FHASSAG", "farheap", CancellationToken.None);

        var drive = await adapter.GetAllDrives(site.Id, CancellationToken.None);
        var driveId = drive[0].Id;
        var folders = await adapter.GetAllFolders(driveId);
        var f = folders.FirstOrDefault(folder => folder.Name.Contains(" payable"));
        var folders1 = await adapter.GetAllFolders(driveId, f.Id);
        var f1 = folders1.FirstOrDefault(folder => folder.Name.Contains("SAG Vendors"));
        var folders2 = await adapter.GetAllFolders(driveId, f1.Id);
        var f2 = folders2.FirstOrDefault(folder => folder.Name == "ONP");
        var folders3 = await adapter.GetAllFolders(driveId, f2.Id);
        var f4 = folders3.FirstOrDefault(folder => folder.Name == "LL Documents");
        _ = JsonSerializer.Serialize(new { DriveId = driveId, FolderId = f4.Id });
    }

    private static ServiceProvider CreateServiceProvider()
    {
        IServiceCollection services = new ServiceCollection();
        var configurationManager = CreateConfiguration();
        services.AddKeyCloak(configurationManager);
        return services.BuildServiceProvider();
    }

    private static ConfigurationManager CreateConfiguration()
    {
        ConfigurationManager configurationManager = new();
        configurationManager.AddJsonFile("appsettings.json");
        configurationManager.AddJsonFile("appsettings.Development.json", optional: true);
        return configurationManager;
    }

    private static string RequireSetting(IConfiguration configuration, string key)
    {
        var value = configuration[key];

        if (string.IsNullOrWhiteSpace(value))
        {
            Assert.Inconclusive($"Set '{key}' in appsettings.Development.json or environment variables to run this integration test.");
        }

        return value!;
    }
}

public sealed class DefaultHttpClientFactory : IHttpClientFactory, IDisposable
{
    public static DefaultHttpClientFactory Instance { get; } = new();

    private readonly Lazy<HttpMessageHandler> _handlerLazy = new(() => new HttpClientHandler());

    public HttpClient CreateClient(string name) => new(_handlerLazy.Value, disposeHandler: false);

    public void Dispose()
    {
        if (_handlerLazy.IsValueCreated)
        {
            _handlerLazy.Value.Dispose();
        }
    }
}
