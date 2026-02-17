namespace DecoSOP.Models;

public class Category
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int SortOrder { get; set; }
    public List<SopDocument> Documents { get; set; } = [];
}
