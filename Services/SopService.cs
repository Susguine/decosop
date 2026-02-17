using DecoSOP.Data;
using DecoSOP.Models;
using Microsoft.EntityFrameworkCore;

namespace DecoSOP.Services;

public class SopService
{
    private readonly AppDbContext _db;

    public SopService(AppDbContext db) => _db = db;

    // --- Categories ---

    public async Task<List<Category>> GetCategoriesAsync()
        => await _db.Categories
            .Include(c => c.Documents.OrderBy(d => d.SortOrder))
            .OrderBy(c => c.SortOrder)
            .ToListAsync();

    public async Task<Category> CreateCategoryAsync(string name)
    {
        var maxSort = await _db.Categories.MaxAsync(c => (int?)c.SortOrder) ?? -1;
        var category = new Category { Name = name.Trim(), SortOrder = maxSort + 1 };
        _db.Categories.Add(category);
        await _db.SaveChangesAsync();
        return category;
    }

    public async Task RenameCategoryAsync(int id, string newName)
    {
        var cat = await _db.Categories.FindAsync(id);
        if (cat is null) return;
        cat.Name = newName.Trim();
        await _db.SaveChangesAsync();
    }

    public async Task DeleteCategoryAsync(int id)
    {
        var cat = await _db.Categories.FindAsync(id);
        if (cat is null) return;
        _db.Categories.Remove(cat);
        await _db.SaveChangesAsync();
    }

    // --- Documents ---

    public async Task<SopDocument?> GetDocumentAsync(int id)
        => await _db.Documents
            .Include(d => d.Category)
            .FirstOrDefaultAsync(d => d.Id == id);

    public async Task<SopDocument> CreateDocumentAsync(int categoryId, string title)
    {
        var maxSort = await _db.Documents
            .Where(d => d.CategoryId == categoryId)
            .MaxAsync(d => (int?)d.SortOrder) ?? -1;

        var doc = new SopDocument
        {
            CategoryId = categoryId,
            Title = title.Trim(),
            MarkdownContent = $"# {title.Trim()}\n\nStart writing your SOP here...",
            SortOrder = maxSort + 1
        };

        _db.Documents.Add(doc);
        await _db.SaveChangesAsync();
        return doc;
    }

    public async Task UpdateDocumentAsync(int id, string title, string markdownContent)
    {
        var doc = await _db.Documents.FindAsync(id);
        if (doc is null) return;
        doc.Title = title.Trim();
        doc.MarkdownContent = markdownContent;
        doc.UpdatedAt = DateTime.UtcNow;
        await _db.SaveChangesAsync();
    }

    public async Task DeleteDocumentAsync(int id)
    {
        var doc = await _db.Documents.FindAsync(id);
        if (doc is null) return;
        _db.Documents.Remove(doc);
        await _db.SaveChangesAsync();
    }

    public async Task<List<SopDocument>> SearchAsync(string query)
    {
        if (string.IsNullOrWhiteSpace(query)) return [];

        var term = query.Trim().ToLower();
        return await _db.Documents
            .Include(d => d.Category)
            .Where(d => d.Title.ToLower().Contains(term)
                     || d.MarkdownContent.ToLower().Contains(term))
            .OrderBy(d => d.Category.SortOrder)
            .ThenBy(d => d.SortOrder)
            .ToListAsync();
    }
}
