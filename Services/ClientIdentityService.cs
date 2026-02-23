namespace DecoSOP.Services;

/// <summary>
/// Scoped service that captures the client's IP address from the initial HTTP request.
/// Used to key per-machine preferences (favorites, pins, folder colors).
/// </summary>
public class ClientIdentityService
{
    public string ClientId { get; }

    public ClientIdentityService(IHttpContextAccessor accessor)
    {
        ClientId = accessor.HttpContext?.Connection.RemoteIpAddress?.ToString() ?? "unknown";
    }
}
