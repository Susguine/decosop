using DecoSOP.Data;
using DecoSOP.Models;
using Microsoft.EntityFrameworkCore;

namespace DecoSOP.Services;

/// <summary>
/// Scoped cache that eliminates duplicate DB calls between the sidebar and pages.
/// Uses a SemaphoreSlim to serialize DB access so Task.Run callers from NavMenu
/// and page components never issue concurrent queries on the same DbContext.
/// Once results are cached, the semaphore is bypassed entirely (instant reads).
/// </summary>
public class DataCacheService
{
    private readonly WebSopService _webSop;
    private readonly SopFileService _sopFile;
    private readonly DocumentService _doc;
    private readonly WebDocService _webDoc;
    private readonly UserPreferenceService _prefs;
    private readonly AppDbContext _db;
    private readonly SemaphoreSlim _sem = new(1, 1);

    public DataCacheService(WebSopService webSop, SopFileService sopFile, DocumentService doc, WebDocService webDoc,
        UserPreferenceService prefs, AppDbContext db)
    {
        _webSop = webSop;
        _sopFile = sopFile;
        _doc = doc;
        _webDoc = webDoc;
        _prefs = prefs;
        _db = db;
    }

    // ---- Sop (files) ----
    private List<SopCategory>? _sopTree;
    private List<SopCategory>? _sopFavCats;
    private List<SopFile>? _sopFavDocs;
    private Dictionary<int, UserPreference>? _sopCatPrefs;
    private Dictionary<int, UserPreference>? _sopFilePrefs;

    public async Task<List<SopCategory>> GetSopTreeAsync()
    {
        if (_sopTree is not null) return _sopTree;
        await _sem.WaitAsync();
        try { return _sopTree ??= await _sopFile.GetCategoryTreeAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<Dictionary<int, UserPreference>> GetSopCategoryPrefsAsync()
    {
        if (_sopCatPrefs is not null) return _sopCatPrefs;
        await _sem.WaitAsync();
        try { return _sopCatPrefs ??= await _prefs.GetAllForTypeAsync(nameof(SopCategory)); }
        finally { _sem.Release(); }
    }

    public async Task<Dictionary<int, UserPreference>> GetSopFilePrefsAsync()
    {
        if (_sopFilePrefs is not null) return _sopFilePrefs;
        await _sem.WaitAsync();
        try { return _sopFilePrefs ??= await _prefs.GetAllForTypeAsync(nameof(SopFile)); }
        finally { _sem.Release(); }
    }

    public async Task<List<SopCategory>> GetSopFavoriteCategoriesAsync()
    {
        if (_sopFavCats is not null) return _sopFavCats;
        await _sem.WaitAsync();
        try
        {
            if (_sopFavCats is not null) return _sopFavCats;
            var favIds = await _prefs.GetFavoritedIdsAsync(nameof(SopCategory));
            if (favIds.Count == 0) { _sopFavCats = []; return _sopFavCats; }
            var favIdSet = new HashSet<int>(favIds);
            var tree = _sopTree ?? await _sopFile.GetCategoryTreeAsync();
            _sopTree ??= tree;
            _sopFavCats = FlattenTree<SopCategory>(tree, c => c.Children)
                .Where(c => favIdSet.Contains(c.Id)).ToList();
            return _sopFavCats;
        }
        finally { _sem.Release(); }
    }

    public async Task<List<SopFile>> GetSopFavoriteDocumentsAsync()
    {
        if (_sopFavDocs is not null) return _sopFavDocs;
        await _sem.WaitAsync();
        try
        {
            if (_sopFavDocs is not null) return _sopFavDocs;
            var favIds = await _prefs.GetFavoritedIdsAsync(nameof(SopFile));
            if (favIds.Count == 0) { _sopFavDocs = []; return _sopFavDocs; }
            _sopFavDocs = await _db.SopFiles
                .Include(d => d.Category)
                .Where(d => favIds.Contains(d.Id))
                .OrderBy(d => d.Title)
                .ToListAsync();
            return _sopFavDocs;
        }
        finally { _sem.Release(); }
    }

    public async Task<SopCategory?> FindSopCategoryAsync(int id)
    {
        var tree = await GetSopTreeAsync();
        return FindInTree<SopCategory>(tree, id, c => c.Id, c => c.Children);
    }

    public async Task<List<(int Id, string Name)>> GetSopBreadcrumbsAsync(int categoryId)
    {
        var cat = await FindSopCategoryAsync(categoryId);
        return BuildBreadcrumbs(cat, c => c.Parent, c => (c.Id, c.Name));
    }

    public void InvalidateSop() { _sopTree = null; _sopFavCats = null; _sopFavDocs = null; _sopCatPrefs = null; _sopFilePrefs = null; }
    public void InvalidateSopFavorites() { _sopFavCats = null; _sopFavDocs = null; _sopCatPrefs = null; _sopFilePrefs = null; }

    // ---- WebSop (HTML editor) ----
    private List<Category>? _webSopTree;
    private List<Category>? _webSopFavCats;
    private List<SopDocument>? _webSopFavDocs;
    private Dictionary<int, UserPreference>? _webSopCatPrefs;
    private Dictionary<int, UserPreference>? _webSopDocPrefs;

    public async Task<List<Category>> GetWebSopTreeAsync()
    {
        if (_webSopTree is not null) return _webSopTree;
        await _sem.WaitAsync();
        try { return _webSopTree ??= await _webSop.GetCategoryTreeAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<Dictionary<int, UserPreference>> GetWebSopCategoryPrefsAsync()
    {
        if (_webSopCatPrefs is not null) return _webSopCatPrefs;
        await _sem.WaitAsync();
        try { return _webSopCatPrefs ??= await _prefs.GetAllForTypeAsync(nameof(Category)); }
        finally { _sem.Release(); }
    }

    public async Task<Dictionary<int, UserPreference>> GetWebSopDocPrefsAsync()
    {
        if (_webSopDocPrefs is not null) return _webSopDocPrefs;
        await _sem.WaitAsync();
        try { return _webSopDocPrefs ??= await _prefs.GetAllForTypeAsync(nameof(SopDocument)); }
        finally { _sem.Release(); }
    }

    public async Task<List<Category>> GetWebSopFavoriteCategoriesAsync()
    {
        if (_webSopFavCats is not null) return _webSopFavCats;
        await _sem.WaitAsync();
        try
        {
            if (_webSopFavCats is not null) return _webSopFavCats;
            var favIds = await _prefs.GetFavoritedIdsAsync(nameof(Category));
            if (favIds.Count == 0) { _webSopFavCats = []; return _webSopFavCats; }
            var favIdSet = new HashSet<int>(favIds);
            var tree = _webSopTree ?? await _webSop.GetCategoryTreeAsync();
            _webSopTree ??= tree;
            _webSopFavCats = FlattenTree<Category>(tree, c => c.Children)
                .Where(c => favIdSet.Contains(c.Id)).ToList();
            return _webSopFavCats;
        }
        finally { _sem.Release(); }
    }

    public async Task<List<SopDocument>> GetWebSopFavoriteDocumentsAsync()
    {
        if (_webSopFavDocs is not null) return _webSopFavDocs;
        await _sem.WaitAsync();
        try
        {
            if (_webSopFavDocs is not null) return _webSopFavDocs;
            var favIds = await _prefs.GetFavoritedIdsAsync(nameof(SopDocument));
            if (favIds.Count == 0) { _webSopFavDocs = []; return _webSopFavDocs; }
            _webSopFavDocs = await _db.Documents
                .Include(d => d.Category)
                .Where(d => favIds.Contains(d.Id))
                .OrderBy(d => d.Title)
                .ToListAsync();
            return _webSopFavDocs;
        }
        finally { _sem.Release(); }
    }

    public async Task<Category?> FindWebSopCategoryAsync(int id)
    {
        var tree = await GetWebSopTreeAsync();
        return FindInTree<Category>(tree, id, c => c.Id, c => c.Children);
    }

    public async Task<List<(int Id, string Name)>> GetWebSopBreadcrumbsAsync(int categoryId)
    {
        var cat = await FindWebSopCategoryAsync(categoryId);
        return BuildBreadcrumbs(cat, c => c.Parent, c => (c.Id, c.Name));
    }

    public void InvalidateWebSop() { _webSopTree = null; _webSopFavCats = null; _webSopFavDocs = null; _webSopCatPrefs = null; _webSopDocPrefs = null; }
    public void InvalidateWebSopFavorites() { _webSopFavCats = null; _webSopFavDocs = null; _webSopCatPrefs = null; _webSopDocPrefs = null; }

    // ---- Document ----
    private List<DocumentCategory>? _docTree;
    private List<DocumentCategory>? _docFavCats;
    private List<OfficeDocument>? _docFavDocs;
    private Dictionary<int, UserPreference>? _docCatPrefs;
    private Dictionary<int, UserPreference>? _docDocPrefs;

    public async Task<List<DocumentCategory>> GetDocTreeAsync()
    {
        if (_docTree is not null) return _docTree;
        await _sem.WaitAsync();
        try { return _docTree ??= await _doc.GetCategoryTreeAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<Dictionary<int, UserPreference>> GetDocCategoryPrefsAsync()
    {
        if (_docCatPrefs is not null) return _docCatPrefs;
        await _sem.WaitAsync();
        try { return _docCatPrefs ??= await _prefs.GetAllForTypeAsync(nameof(DocumentCategory)); }
        finally { _sem.Release(); }
    }

    public async Task<Dictionary<int, UserPreference>> GetDocDocPrefsAsync()
    {
        if (_docDocPrefs is not null) return _docDocPrefs;
        await _sem.WaitAsync();
        try { return _docDocPrefs ??= await _prefs.GetAllForTypeAsync(nameof(OfficeDocument)); }
        finally { _sem.Release(); }
    }

    public async Task<List<DocumentCategory>> GetDocFavoriteCategoriesAsync()
    {
        if (_docFavCats is not null) return _docFavCats;
        await _sem.WaitAsync();
        try
        {
            if (_docFavCats is not null) return _docFavCats;
            var favIds = await _prefs.GetFavoritedIdsAsync(nameof(DocumentCategory));
            if (favIds.Count == 0) { _docFavCats = []; return _docFavCats; }
            var favIdSet = new HashSet<int>(favIds);
            var tree = _docTree ?? await _doc.GetCategoryTreeAsync();
            _docTree ??= tree;
            _docFavCats = FlattenTree<DocumentCategory>(tree, c => c.Children)
                .Where(c => favIdSet.Contains(c.Id)).ToList();
            return _docFavCats;
        }
        finally { _sem.Release(); }
    }

    public async Task<List<OfficeDocument>> GetDocFavoriteDocumentsAsync()
    {
        if (_docFavDocs is not null) return _docFavDocs;
        await _sem.WaitAsync();
        try
        {
            if (_docFavDocs is not null) return _docFavDocs;
            var favIds = await _prefs.GetFavoritedIdsAsync(nameof(OfficeDocument));
            if (favIds.Count == 0) { _docFavDocs = []; return _docFavDocs; }
            _docFavDocs = await _db.OfficeDocuments
                .Include(d => d.Category)
                .Where(d => favIds.Contains(d.Id))
                .OrderBy(d => d.Title)
                .ToListAsync();
            return _docFavDocs;
        }
        finally { _sem.Release(); }
    }

    public async Task<DocumentCategory?> FindDocCategoryAsync(int id)
    {
        var tree = await GetDocTreeAsync();
        return FindInTree<DocumentCategory>(tree, id, c => c.Id, c => c.Children);
    }

    public async Task<List<(int Id, string Name)>> GetDocBreadcrumbsAsync(int categoryId)
    {
        var cat = await FindDocCategoryAsync(categoryId);
        return BuildBreadcrumbs(cat, c => c.Parent, c => (c.Id, c.Name));
    }

    public void InvalidateDoc() { _docTree = null; _docFavCats = null; _docFavDocs = null; _docCatPrefs = null; _docDocPrefs = null; }
    public void InvalidateDocFavorites() { _docFavCats = null; _docFavDocs = null; _docCatPrefs = null; _docDocPrefs = null; }

    // ---- WebDoc ----
    private List<WebDocCategory>? _webDocTree;
    private List<WebDocCategory>? _webDocFavCats;
    private List<WebDocument>? _webDocFavDocs;
    private Dictionary<int, UserPreference>? _webDocCatPrefs;
    private Dictionary<int, UserPreference>? _webDocDocPrefs;

    public async Task<List<WebDocCategory>> GetWebDocTreeAsync()
    {
        if (_webDocTree is not null) return _webDocTree;
        await _sem.WaitAsync();
        try { return _webDocTree ??= await _webDoc.GetCategoryTreeAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<Dictionary<int, UserPreference>> GetWebDocCategoryPrefsAsync()
    {
        if (_webDocCatPrefs is not null) return _webDocCatPrefs;
        await _sem.WaitAsync();
        try { return _webDocCatPrefs ??= await _prefs.GetAllForTypeAsync(nameof(WebDocCategory)); }
        finally { _sem.Release(); }
    }

    public async Task<Dictionary<int, UserPreference>> GetWebDocDocPrefsAsync()
    {
        if (_webDocDocPrefs is not null) return _webDocDocPrefs;
        await _sem.WaitAsync();
        try { return _webDocDocPrefs ??= await _prefs.GetAllForTypeAsync(nameof(WebDocument)); }
        finally { _sem.Release(); }
    }

    public async Task<List<WebDocCategory>> GetWebDocFavoriteCategoriesAsync()
    {
        if (_webDocFavCats is not null) return _webDocFavCats;
        await _sem.WaitAsync();
        try
        {
            if (_webDocFavCats is not null) return _webDocFavCats;
            var favIds = await _prefs.GetFavoritedIdsAsync(nameof(WebDocCategory));
            if (favIds.Count == 0) { _webDocFavCats = []; return _webDocFavCats; }
            var favIdSet = new HashSet<int>(favIds);
            var tree = _webDocTree ?? await _webDoc.GetCategoryTreeAsync();
            _webDocTree ??= tree;
            _webDocFavCats = FlattenTree<WebDocCategory>(tree, c => c.Children)
                .Where(c => favIdSet.Contains(c.Id)).ToList();
            return _webDocFavCats;
        }
        finally { _sem.Release(); }
    }

    public async Task<List<WebDocument>> GetWebDocFavoriteDocumentsAsync()
    {
        if (_webDocFavDocs is not null) return _webDocFavDocs;
        await _sem.WaitAsync();
        try
        {
            if (_webDocFavDocs is not null) return _webDocFavDocs;
            var favIds = await _prefs.GetFavoritedIdsAsync(nameof(WebDocument));
            if (favIds.Count == 0) { _webDocFavDocs = []; return _webDocFavDocs; }
            _webDocFavDocs = await _db.WebDocuments
                .Include(d => d.Category)
                .Where(d => favIds.Contains(d.Id))
                .OrderBy(d => d.Title)
                .ToListAsync();
            return _webDocFavDocs;
        }
        finally { _sem.Release(); }
    }

    public async Task<WebDocCategory?> FindWebDocCategoryAsync(int id)
    {
        var tree = await GetWebDocTreeAsync();
        return FindInTree<WebDocCategory>(tree, id, c => c.Id, c => c.Children);
    }

    public async Task<List<(int Id, string Name)>> GetWebDocBreadcrumbsAsync(int categoryId)
    {
        var cat = await FindWebDocCategoryAsync(categoryId);
        return BuildBreadcrumbs(cat, c => c.Parent, c => (c.Id, c.Name));
    }

    public void InvalidateWebDoc() { _webDocTree = null; _webDocFavCats = null; _webDocFavDocs = null; _webDocCatPrefs = null; _webDocDocPrefs = null; }
    public void InvalidateWebDocFavorites() { _webDocFavCats = null; _webDocFavDocs = null; _webDocCatPrefs = null; _webDocDocPrefs = null; }

    // ---- Invalidate all ----
    public void InvalidateAll() { InvalidateSop(); InvalidateWebSop(); InvalidateDoc(); InvalidateWebDoc(); }

    /// <summary>
    /// Runs an async action under the shared semaphore so it never overlaps with
    /// cache reads or other DB operations on the same scoped DbContext.
    /// Use this for direct service calls (toggle favorite, create category, etc.)
    /// that bypass the cache.
    /// </summary>
    public async Task RunExclusiveAsync(Func<Task> action)
    {
        await _sem.WaitAsync();
        try { await action(); }
        finally { _sem.Release(); }
    }

    public async Task<T> RunExclusiveAsync<T>(Func<Task<T>> action)
    {
        await _sem.WaitAsync();
        try { return await action(); }
        finally { _sem.Release(); }
    }

    // ---- Helpers ----

    private static T? FindInTree<T>(IEnumerable<T> roots, int id,
        Func<T, int> getId, Func<T, IEnumerable<T>> getChildren) where T : class
    {
        foreach (var node in roots)
        {
            if (getId(node) == id) return node;
            var found = FindInTree(getChildren(node), id, getId, getChildren);
            if (found is not null) return found;
        }
        return null;
    }

    private static List<T> FlattenTree<T>(IEnumerable<T> roots, Func<T, IEnumerable<T>> getChildren)
    {
        var result = new List<T>();
        foreach (var node in roots)
        {
            result.Add(node);
            result.AddRange(FlattenTree(getChildren(node), getChildren));
        }
        return result;
    }

    private static List<(int Id, string Name)> BuildBreadcrumbs<T>(
        T? current, Func<T, T?> getParent, Func<T, (int Id, string Name)> toTuple) where T : class
    {
        var crumbs = new List<(int Id, string Name)>();
        while (current is not null)
        {
            crumbs.Insert(0, toTuple(current));
            current = getParent(current);
        }
        return crumbs;
    }
}
