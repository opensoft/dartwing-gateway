using Microsoft.Extensions.Caching.Memory;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace DartWing.Microsoft;

public sealed class GraphApiAdapter : IDisposable
{
    private readonly string _email;
    private readonly string _tenantId;
    private readonly string _tokenKey;
    private readonly IHttpClientFactory? _clientFactory;
    private readonly IMemoryCache _memoryCache;
    private readonly GraphServiceClient _graphClient;

    public GraphApiAdapter(string token, IHttpClientFactory? clientFactory, IMemoryCache memoryCache)
    {
        var (email, tenantId) = MicrosoftExtension.GetEmailAndTenantId(token);
        _email = email;
        _tenantId = tenantId;
        _tokenKey = $"{tenantId}__{email}";
        _clientFactory = clientFactory;
        _memoryCache = memoryCache;
        _graphClient = new GraphServiceClient(new CustomAuthProvider(token));
    }

    public async Task<User?> Me(CancellationToken ct)
    {
        var me = await _graphClient.Me.GetAsync(cancellationToken: ct);
        return me;
    }
    
    public async Task<MicrosoftDriveInfo[]?> Drives(CancellationToken ct)
    {
        var drivesKey = "Microsoft:Drives:" + _tokenKey;
        if (_memoryCache.TryGetValue(drivesKey, out MicrosoftDriveInfo[]? drs) && drs != null) return drs;
        
        var drives = await _graphClient.Drives.GetAsync(cancellationToken: ct);
        var result = drives?.Value?.Select(s => new MicrosoftDriveInfo(s.Id, s.Name)).ToArray();
        if (result != null) _memoryCache.Set(drivesKey, result, TimeSpan.FromSeconds(12));
        return result;
    }

    public async Task<List<MicrosoftFolderInfo>> GetMyFolders(CancellationToken ct)
    {
        var meDrive = await _graphClient.Me.Drive.GetAsync(cancellationToken: ct);
        var allFolders = await GetAllFolders(meDrive!.Id!, ct: ct);
        return allFolders;
    }


    public async Task<Site?> GetRootSite(CancellationToken ct)
    {
        var site = await _graphClient.Sites["root"].GetAsync(cancellationToken: ct);
        return site;
    }


    public async Task<SiteCollectionResponse?> GetSites(CancellationToken ct)
    {
        var sitesKey = "Microsoft:Sites:" + _tokenKey;
        if (_memoryCache.TryGetValue(sitesKey, out SiteCollectionResponse? sts) && sts != null)
        {
           return sts;
        }
        var result = await _graphClient.Sites
            .WithUrl("https://graph.microsoft.com/v1.0/sites?search=LastModifiedTime>=2015-01-01")
            .GetAsync(cancellationToken: ct);

        if (result?.Value != null && result.Value.Count > 0)
            _memoryCache.Set(sitesKey, result, TimeSpan.FromSeconds(12));
        
        return result;
    }

    private async Task<List<Site>> FetchAllSites(CancellationToken ct)
    {
        var cacheKey = $"Microsoft:Sites:All__{_tokenKey}";
        if (_memoryCache.TryGetValue(cacheKey, out List<Site>? cached) && cached != null)
            return cached;

        List<Site> allSites = [];
        var request = await _graphClient.Sites
            .WithUrl("https://graph.microsoft.com/v1.0/sites?search=LastModifiedTime>=2015-01-01")
            .GetAsync(cancellationToken: ct);
        while (request?.Value != null)
        {
            allSites.AddRange(request.Value);
            if (request.OdataNextLink != null)
                request = await _graphClient.Sites.WithUrl(request.OdataNextLink)
                    .GetAsync(cancellationToken: ct);
            else
                break;
        }

        if (allSites.Count > 0)
            _memoryCache.Set(cacheKey, allSites, TimeSpan.FromSeconds(30));

        return allSites;
    }

    private static MicrosoftSiteInfo ToSiteInfo(Site s) => new(s.Id, s.Name)
    {
        DisplayName = s.DisplayName,
        Description = s.Description,
        LastModifiedDateTime = s.LastModifiedDateTime
    };

    public async Task<MicrosoftSiteInfo[]> GetAllSites(CancellationToken ct)
    {
        var allSites = await FetchAllSites(ct);
        return allSites.Select(s => new MicrosoftSiteInfo(s.Id, s.Name)).ToArray();
    }

    public async Task<MicrosoftSiteInfo[]> SearchSites(string query, CancellationToken ct)
    {
        try
        {
            var allSites = await FetchAllSites(ct);
            // Graph search is unreliable for partial strings — filter locally instead
            return allSites
                .Where(s => !s.Id?.Contains("-my.sharepoint.com,") == true
                            && !s.WebUrl?.Contains("-my.sharepoint.com") == true)
                .Where(s => s.Name?.Contains(query, StringComparison.OrdinalIgnoreCase) == true
                            || s.DisplayName?.Contains(query, StringComparison.OrdinalIgnoreCase) == true)
                .Select(ToSiteInfo)
                .ToArray();
        }
        catch (ODataError)
        {
            return [];
        }
    }

    public async Task<MicrosoftSiteInfo?> GetSite(string siteName, string? tenantName, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(tenantName))
            tenantName = await GetTenantName(ct);

        if (siteName == $"{tenantName}.sharepoint.com")
        {
            var allSites = await GetSites(ct);
            var site = allSites.Value.FirstOrDefault(s => s.Name == siteName);
            return site == null ? null : new MicrosoftSiteInfo(site.Id, site.Name);
        }
        
        var s = await _graphClient.Sites[$"{tenantName}.sharepoint.com:/sites/{siteName}:/"].GetAsync(cancellationToken: ct);
        return s == null ? null : new MicrosoftSiteInfo(s.Id, s.Name);
    }

    public async Task<string?> GetTenantName(CancellationToken ct)
    {
        try
        {
            var site = await GetRootSite(ct);

            return site?.WebUrl?.Substring("https://".Length, site.WebUrl.IndexOf('.') - "https://".Length);
        }
        catch (Exception e)
        {
            return null;
        }
    }

    public async Task<List<MicrosoftSiteInfo>> GetAllSitesWithDrives(CancellationToken ct)
    {
        SiteCollectionResponse? sites = null;
        var sitesKey = "Microsoft:Sites:" + _tokenKey;
        if (_memoryCache.TryGetValue(sitesKey, out SiteCollectionResponse? sts) && sts != null)
        {
            sites = sts;
        }
        else
        {
            sites = await GetSites(ct);
            _memoryCache.Set(sitesKey, sites, TimeSpan.FromSeconds(24));
        }
        if (sites?.Value == null) return [];
        List<Task<MicrosoftSiteInfo?>> tasks = [];
        foreach (var s in sites.Value.Where(s => !s.Id.Contains("-my.sharepoint.com,")))
        {
            var driveTask = GetDrivesPrivate(s, ct);
#if DEBUG
            await driveTask;
#endif
            tasks.Add(driveTask);
            if (tasks.Count(x => !x.IsCompleted) > 6)
            {
                await Task.WhenAny(tasks.Where(x => !x.IsCompleted));
            }
        }
        
        await Task.WhenAll(tasks);

        return tasks.Select(x => x.Result)
            .Where(x => x?.Drives != null && x.Drives.Count > 0 && x.Drives[0].Folders != null).ToList();
    }

    private async Task<MicrosoftSiteInfo?> GetDrivesPrivate(Site s, CancellationToken ct)
    {
        var sitesKey = $"Microsoft:Sites:{s.Id}__{_tokenKey}";
        if (_memoryCache.TryGetValue(sitesKey, out MicrosoftSiteInfo? sts) && sts != null)
        {
            return sts.Drives.Count == 0 ? null : sts;
        }
        
        try
        {
            var siteDrives = await _graphClient.Sites[s.Id].Drives.GetAsync(cancellationToken: ct);
            if (siteDrives?.Value == null || siteDrives.Value.Count == 0) return null;
            var st = new MicrosoftSiteInfo(s.Id, s.Name)
            {
                DisplayName = s.DisplayName, Description = s.Description, LastModifiedDateTime = s.LastModifiedDateTime
            };
            foreach (var d in siteDrives.Value)
            {
                if (d == null) continue;
                var fld = await GetAllFolders(d.Id!, ct: ct);
                //if (fld.Count == 0) continue;
                st.Drives.Add(new MicrosoftDriveInfo(d.Id, d.Name) {Folders = fld, LastModifiedDateTime = d.LastModifiedDateTime});
            }
            _memoryCache.Set(sitesKey, st, TimeSpan.FromSeconds(12));
            return st.Drives.Count == 0 ? null : st;
        }
        catch (Exception e)
        {
            _memoryCache.Set(sitesKey, new MicrosoftSiteInfo(s.Id, s.Name), TimeSpan.FromSeconds(24));
            return null;
        }
    }

    public async Task<MicrosoftDriveInfo[]> GetAllDrives(string siteId, CancellationToken ct)
    {
        try
        {
            var request = await _graphClient.Sites[siteId].Drives.GetAsync(cancellationToken: ct);
            List<Drive> allDrives = [];
            while (request?.Value != null)
            {
                allDrives.AddRange(request.Value);
                if (request.OdataNextLink != null)
                    request = await _graphClient.Sites[siteId].Drives
                        .WithUrl(request.OdataNextLink).GetAsync(cancellationToken: ct);
                else
                    break;
            }
            return allDrives.Select(s => new MicrosoftDriveInfo(s.Id, s.Name)
                { LastModifiedDateTime = s.LastModifiedDateTime }).ToArray();
        }
        catch (ODataError)
        {
            return [];
        }
    }

    public async Task<MicrosoftDriveInfo[]?> GetAllDrives(CancellationToken ct)
    {
        var drives = await _graphClient.Drives.GetAsync(cancellationToken: ct);
        return drives?.Value?.Select(s => new MicrosoftDriveInfo(s.Id, s.Name)).ToArray();
    }

    public async Task<List<MicrosoftFolderInfo>> GetAllFolders(string driveId, string itemId = "root",
        bool recursive = false, CancellationToken ct = default)
    {
        var cacheKey = $"Microsoft:Drives:{_tokenKey}:Folders:{driveId}:{itemId}:{recursive}";
        if (_memoryCache.TryGetValue(cacheKey, out List<MicrosoftFolderInfo>? cached) && cached != null)
            return cached;

        List<MicrosoftFolderInfo> folders = [];
        await FetchFoldersRecursive(driveId, itemId, null, folders, recursive, ct);

        _memoryCache.Set(cacheKey, folders, TimeSpan.FromSeconds(12));
        return folders;
    }
    
    public async Task<string?> GetDriveName(string driveId, CancellationToken ct)
    {
        try
        {
            var drive = await _graphClient.Drives[driveId].GetAsync(cancellationToken: ct);
            return drive?.Name;
        }
        catch (ODataError)
        {
            return null;
        }
    }

    public async Task<string?> GetFolderName(string driveId, string folderId, CancellationToken ct)
    {
        try
        {
            var item = await _graphClient.Drives[driveId].Items[folderId].GetAsync(cancellationToken: ct);
            return item?.Name;
        }
        catch (ODataError)
        {
            return null;
        }
    }

    public async Task<bool> IsFileExists(string driveId, string itemId, string fileName, CancellationToken ct)
    {
        try
        {
            await _graphClient.Drives[driveId].Items[itemId].ItemWithPath(fileName).GetAsync(cancellationToken: ct);

            return true;
        }
        catch (Exception)
        {
            // ignored
        }

        return false;
    }

    private async Task FetchFoldersRecursive(string driveId, string itemId, string? parentFolderId,
        List<MicrosoftFolderInfo> folders, bool recursive, CancellationToken ct)
    {
        try
        {
            var children = await _graphClient.Drives[driveId].Items[itemId].Children.GetAsync(cancellationToken: ct);

            while (children?.Value != null)
            {
                foreach (var item in children.Value)
                {
                    if (item.Folder == null) continue;
                    var folderInfo = new MicrosoftFolderInfo
                    {
                        Id = item.Id!,
                        Name = item.Name!,
                        ParentFolderId = parentFolderId,
                        DriveId = driveId,
                        LastModifiedDateTime = item.LastModifiedDateTime
                    };

                    folders.Add(folderInfo);

                    if (recursive) await FetchFoldersRecursive(driveId, item.Id!, item.Id!, folders, true, ct);
                }

                if (children.OdataNextLink != null)
                    children = await _graphClient.Drives[driveId].Items[itemId].Children
                        .WithUrl(children.OdataNextLink).GetAsync(cancellationToken: ct);
                else
                    break;
            }
        }
        catch (ODataError)
        {
        }
    }
    
    public async Task<DriveItem?> UploadFile(string driveId, string itemId, string fileName, string fileContentType, Stream stream, CancellationToken ct)
    {
        try
        {
            var file = await _graphClient.Drives[driveId].Items[itemId].ItemWithPath(fileName).Content.PutAsync(stream, cancellationToken: ct);

            return file;
        }
        catch (ODataError)
        {
        }

        return null;
    }

    public async Task<DriveItem?> UploadFileByLink(string driveId, string folderId, string fileUrl,
        CancellationToken ct)
    {
        using var message = await _clientFactory!.CreateClient().GetAsync(fileUrl, ct);

        //var contentType = message.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";
        var originalFileName = message.Content.Headers.ContentDisposition?.FileName;
        if (!string.IsNullOrEmpty(originalFileName)) originalFileName = Uri.UnescapeDataString(originalFileName);
        var fileName = originalFileName;
        if (string.IsNullOrEmpty(fileName))
            fileName = $"{Path.GetFileNameWithoutExtension(fileUrl)}_{Random.Shared.Next(9999)}{Path.GetExtension(fileUrl)}";

        if (!string.IsNullOrEmpty(fileName) &&
            await IsFileExists(driveId, folderId, fileName, ct))
        {
            fileName =
                $"{Path.GetFileNameWithoutExtension(fileName)}_{Random.Shared.Next(99999)}{Path.GetExtension(fileName)}";
        }

        await using var fileStream = await message.Content.ReadAsStreamAsync(ct);
        try
        {
            var file = await _graphClient.Drives[driveId].Items[folderId].ItemWithPath(fileName).Content
                .PutAsync(fileStream, cancellationToken: ct);

            return file;
        }
        catch (ODataError e)
        {
        }

        return null;
    }

    public sealed class MicrosoftFolderInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? ParentFolderId { get; set; }
        public string DriveId { get; set; }
        public DateTimeOffset? LastModifiedDateTime { get; set; }
    }

    public sealed class MicrosoftDriveInfo
    {
        public MicrosoftDriveInfo(string? id, string? name)
        {
            Id = id;
            Name = name;
        }

        public string? Id { get; set; }
        public string? Name { get; set; }
        public List<MicrosoftFolderInfo> Folders { get; set; } = [];
        public MicrosoftSiteInfo Site { get; set; }
        public DateTimeOffset? LastModifiedDateTime { get; set; }
    }

    public sealed class MicrosoftSiteInfo
    {
        public MicrosoftSiteInfo(string? id, string? name)
        {
            Id = id;
            Name = name;
        }

        public string? Id { get; set; }
        public string? Name { get; set; }
        public List<MicrosoftDriveInfo> Drives { get; set; } = [];
        public DateTimeOffset? LastModifiedDateTime { get; set; }
        public string? Description { get; set; }
        public string? DisplayName { get; set; }
    }

    public void Dispose()
    {
        _graphClient.Dispose();
    }


}
