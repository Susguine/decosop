using System.Reflection;
using System.Text.Json;

namespace DecoSOP.Services;

/// <summary>
/// Singleton service that periodically checks GitHub Releases for a newer version.
/// </summary>
public sealed class UpdateService : IDisposable
{
    private readonly ILogger<UpdateService> _logger;
    private readonly HttpClient _http;
    private readonly Timer _timer;
    private readonly string _currentVersion;

    // Configurable via update-config.json next to the exe
    private bool _enabled = true;
    private string _repoOwner = "Susguine";
    private string _repoName = "DecoSOP";
    private TimeSpan _checkInterval = TimeSpan.FromHours(24);
    private string? _skippedVersion;

    public string? NewVersion { get; private set; }
    public string? ReleaseUrl { get; private set; }
    public string? DownloadUrl { get; private set; }
    public string? ReleaseNotes { get; private set; }
    public bool UpdateAvailable => NewVersion is not null;

    public event Action? OnUpdateChecked;

    public UpdateService(ILogger<UpdateService> logger)
    {
        _logger = logger;
        _http = new HttpClient();
        _http.DefaultRequestHeaders.UserAgent.ParseAdd("DecoSOP-UpdateChecker/1.0");

        _currentVersion = Assembly.GetExecutingAssembly()
            .GetName().Version?.ToString(3) ?? "0.0.0";

        LoadConfig();

        // Initial check after 30 seconds, then every _checkInterval
        _timer = new Timer(async _ => await CheckForUpdateAsync(),
            null,
            _enabled ? TimeSpan.FromSeconds(30) : Timeout.InfiniteTimeSpan,
            _enabled ? _checkInterval : Timeout.InfiniteTimeSpan);
    }

    public async Task CheckForUpdateAsync()
    {
        if (!_enabled) return;

        try
        {
            var url = $"https://api.github.com/repos/{_repoOwner}/{_repoName}/releases/latest";
            var response = await _http.GetAsync(url);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogDebug("Update check returned {Status}", response.StatusCode);
                return;
            }

            var json = await response.Content.ReadFromJsonAsync<JsonElement>();

            var tagName = json.GetProperty("tag_name").GetString()?.TrimStart('v') ?? "";
            var htmlUrl = json.GetProperty("html_url").GetString() ?? "";
            var body = json.TryGetProperty("body", out var bodyProp) ? bodyProp.GetString() ?? "" : "";

            // Find the installer asset or the zip asset
            string? assetUrl = null;
            if (json.TryGetProperty("assets", out var assets))
            {
                foreach (var asset in assets.EnumerateArray())
                {
                    var name = asset.GetProperty("name").GetString() ?? "";
                    if (name.Contains("Setup", StringComparison.OrdinalIgnoreCase) && name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    {
                        assetUrl = asset.GetProperty("browser_download_url").GetString();
                        break;
                    }
                }
                // Fallback to zip if no installer found
                if (assetUrl is null)
                {
                    foreach (var asset in assets.EnumerateArray())
                    {
                        var name = asset.GetProperty("name").GetString() ?? "";
                        if (name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                        {
                            assetUrl = asset.GetProperty("browser_download_url").GetString();
                            break;
                        }
                    }
                }
            }

            if (!Version.TryParse(tagName, out var remoteVer) ||
                !Version.TryParse(_currentVersion, out var localVer))
            {
                _logger.LogDebug("Could not parse versions: remote={Remote}, local={Local}", tagName, _currentVersion);
                return;
            }

            if (remoteVer > localVer && tagName != _skippedVersion)
            {
                NewVersion = tagName;
                ReleaseUrl = htmlUrl;
                DownloadUrl = assetUrl ?? htmlUrl;
                ReleaseNotes = body.Length > 500 ? body[..500] + "..." : body;
                _logger.LogInformation("Update available: {Version}", tagName);
            }
            else
            {
                NewVersion = null;
                ReleaseUrl = null;
                DownloadUrl = null;
                ReleaseNotes = null;
            }

            OnUpdateChecked?.Invoke();
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Update check failed");
        }
    }

    public void SkipVersion(string version)
    {
        _skippedVersion = version;
        NewVersion = null;
        ReleaseUrl = null;
        DownloadUrl = null;
        ReleaseNotes = null;
        SaveConfig();
        OnUpdateChecked?.Invoke();
    }

    public void Dismiss()
    {
        NewVersion = null;
        ReleaseUrl = null;
        DownloadUrl = null;
        ReleaseNotes = null;
        OnUpdateChecked?.Invoke();
    }

    private void LoadConfig()
    {
        try
        {
            var configPath = Path.Combine(AppContext.BaseDirectory, "update-config.json");
            if (!File.Exists(configPath)) return;

            var json = JsonSerializer.Deserialize<JsonElement>(File.ReadAllText(configPath));

            if (json.TryGetProperty("enabled", out var enabled))
                _enabled = enabled.GetBoolean();
            if (json.TryGetProperty("repoOwner", out var owner))
                _repoOwner = owner.GetString() ?? _repoOwner;
            if (json.TryGetProperty("repoName", out var name))
                _repoName = name.GetString() ?? _repoName;
            if (json.TryGetProperty("checkIntervalHours", out var hours))
                _checkInterval = TimeSpan.FromHours(hours.GetInt32());
            if (json.TryGetProperty("skippedVersion", out var skipped))
                _skippedVersion = skipped.GetString();
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to load update config");
        }
    }

    private void SaveConfig()
    {
        try
        {
            var configPath = Path.Combine(AppContext.BaseDirectory, "update-config.json");
            var config = new
            {
                enabled = _enabled,
                repoOwner = _repoOwner,
                repoName = _repoName,
                checkIntervalHours = (int)_checkInterval.TotalHours,
                skippedVersion = _skippedVersion
            };
            File.WriteAllText(configPath, JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to save update config");
        }
    }

    public void Dispose()
    {
        _timer.Dispose();
        _http.Dispose();
    }
}
