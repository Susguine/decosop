namespace DecoSOP.Models;

public class WebDocCategory
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int SortOrder { get; set; }
    public int? ParentId { get; set; }
    public WebDocCategory? Parent { get; set; }
    public List<WebDocCategory> Children { get; set; } = [];
    public List<WebDocument> Documents { get; set; } = [];
}
