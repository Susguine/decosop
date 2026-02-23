namespace DecoSOP.Models;

public class UserPreference
{
    public int Id { get; set; }
    public string ClientId { get; set; } = "";
    public string EntityType { get; set; } = "";
    public int EntityId { get; set; }
    public bool IsFavorited { get; set; }
    public bool IsPinned { get; set; }
    public string? Color { get; set; }
}
