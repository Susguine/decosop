using DecoSOP.Models;

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
    private readonly SemaphoreSlim _sem = new(1, 1);

    public DataCacheService(WebSopService webSop, SopFileService sopFile, DocumentService doc, WebDocService webDoc)
    {
        _webSop = webSop;
        _sopFile = sopFile;
        _doc = doc;
        _webDoc = webDoc;
    }

    // ---- Sop (files) ----
    private List<SopCategory>? _sopTree;
    private List<SopCategory>? _sopFavCats;
    private List<SopFile>? _sopFavDocs;

    public async Task<List<SopCategory>> GetSopTreeAsync()
    {
        if (_sopTree is not null) return _sopTree;
        await _sem.WaitAsync();
        try { return _sopTree ??= await _sopFile.GetCategoryTreeAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<List<SopCategory>> GetSopFavoriteCategoriesAsync()
    {
        if (_sopFavCats is not null) return _sopFavCats;
        await _sem.WaitAsync();
        try { return _sopFavCats ??= await _sopFile.GetFavoriteCategoriesAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<List<SopFile>> GetSopFavoriteDocumentsAsync()
    {
        if (_sopFavDocs is not null) return _sopFavDocs;
        await _sem.WaitAsync();
        try { return _sopFavDocs ??= await _sopFile.GetFavoriteDocumentsAsync(); }
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

    public void InvalidateSop() { _sopTree = null; _sopFavCats = null; _sopFavDocs = null; }
    public void InvalidateSopFavorites() { _sopFavCats = null; _sopFavDocs = null; }

    // ---- WebSop (HTML editor) ----
    private List<Category>? _webSopTree;
    private List<Category>? _webSopFavCats;
    private List<SopDocument>? _webSopFavDocs;

    public async Task<List<Category>> GetWebSopTreeAsync()
    {
        if (_webSopTree is not null) return _webSopTree;
        await _sem.WaitAsync();
        try { return _webSopTree ??= await _webSop.GetCategoryTreeAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<List<Category>> GetWebSopFavoriteCategoriesAsync()
    {
        if (_webSopFavCats is not null) return _webSopFavCats;
        await _sem.WaitAsync();
        try { return _webSopFavCats ??= await _webSop.GetFavoriteCategoriesAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<List<SopDocument>> GetWebSopFavoriteDocumentsAsync()
    {
        if (_webSopFavDocs is not null) return _webSopFavDocs;
        await _sem.WaitAsync();
        try { return _webSopFavDocs ??= await _webSop.GetFavoriteDocumentsAsync(); }
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

    public void InvalidateWebSop() { _webSopTree = null; _webSopFavCats = null; _webSopFavDocs = null; }
    public void InvalidateWebSopFavorites() { _webSopFavCats = null; _webSopFavDocs = null; }

    // ---- Document ----
    private List<DocumentCategory>? _docTree;
    private List<DocumentCategory>? _docFavCats;
    private List<OfficeDocument>? _docFavDocs;

    public async Task<List<DocumentCategory>> GetDocTreeAsync()
    {
        if (_docTree is not null) return _docTree;
        await _sem.WaitAsync();
        try { return _docTree ??= await _doc.GetCategoryTreeAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<List<DocumentCategory>> GetDocFavoriteCategoriesAsync()
    {
        if (_docFavCats is not null) return _docFavCats;
        await _sem.WaitAsync();
        try { return _docFavCats ??= await _doc.GetFavoriteCategoriesAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<List<OfficeDocument>> GetDocFavoriteDocumentsAsync()
    {
        if (_docFavDocs is not null) return _docFavDocs;
        await _sem.WaitAsync();
        try { return _docFavDocs ??= await _doc.GetFavoriteDocumentsAsync(); }
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

    public void InvalidateDoc() { _docTree = null; _docFavCats = null; _docFavDocs = null; }
    public void InvalidateDocFavorites() { _docFavCats = null; _docFavDocs = null; }

    // ---- WebDoc ----
    private List<WebDocCategory>? _webDocTree;
    private List<WebDocCategory>? _webDocFavCats;
    private List<WebDocument>? _webDocFavDocs;

    public async Task<List<WebDocCategory>> GetWebDocTreeAsync()
    {
        if (_webDocTree is not null) return _webDocTree;
        await _sem.WaitAsync();
        try { return _webDocTree ??= await _webDoc.GetCategoryTreeAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<List<WebDocCategory>> GetWebDocFavoriteCategoriesAsync()
    {
        if (_webDocFavCats is not null) return _webDocFavCats;
        await _sem.WaitAsync();
        try { return _webDocFavCats ??= await _webDoc.GetFavoriteCategoriesAsync(); }
        finally { _sem.Release(); }
    }

    public async Task<List<WebDocument>> GetWebDocFavoriteDocumentsAsync()
    {
        if (_webDocFavDocs is not null) return _webDocFavDocs;
        await _sem.WaitAsync();
        try { return _webDocFavDocs ??= await _webDoc.GetFavoriteDocumentsAsync(); }
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

    public void InvalidateWebDoc() { _webDocTree = null; _webDocFavCats = null; _webDocFavDocs = null; }
    public void InvalidateWebDocFavorites() { _webDocFavCats = null; _webDocFavDocs = null; }

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
