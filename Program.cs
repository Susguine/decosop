using DecoSOP.Components;
using DecoSOP.Data;
using DecoSOP.Services;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

builder.Host.UseWindowsService();

// Read port from port.config if it exists (written by installer), otherwise default to 5098
var port = "5098";
var portConfigPath = Path.Combine(AppContext.BaseDirectory, "port.config");
if (File.Exists(portConfigPath))
{
    foreach (var line in File.ReadAllLines(portConfigPath))
    {
        if (line.StartsWith("PORT=", StringComparison.OrdinalIgnoreCase))
        {
            var value = line["PORT=".Length..].Trim();
            if (int.TryParse(value, out var p) && p > 0 && p <= 65535)
                port = value;
            break;
        }
    }
}
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// In development, use project root so import scripts and the app share the same DB.
// In production (Windows Service), use the exe directory.
var dataDir = builder.Environment.IsDevelopment()
    ? builder.Environment.ContentRootPath
    : AppContext.BaseDirectory;
var dbPath = Path.Combine(dataDir, "decosop.db");
DocumentService.DataDirectory = dataDir;
SopFileService.DataDirectory = dataDir;
builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseSqlite($"Data Source={dbPath}"));
builder.Services.AddDbContextFactory<AppDbContext>(options =>
    options.UseSqlite($"Data Source={dbPath}"), ServiceLifetime.Scoped);

builder.Services.AddScoped<WebSopService>();
builder.Services.AddScoped<SopFileService>();
builder.Services.AddScoped<DocumentService>();
builder.Services.AddScoped<WebDocService>();
builder.Services.AddScoped<DataCacheService>();
builder.Services.AddScoped<ContextMenuState>();
builder.Services.AddSingleton<UpdateService>();

var app = builder.Build();

// Auto-create/migrate the database on startup
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();

    // Schema migrations for existing databases (EnsureCreated only creates new DBs)
    var conn = db.Database.GetDbConnection();
    await conn.OpenAsync();
    try
    {
      try
      {
        // Add IsFavorited columns if missing
        foreach (var table in new[] { "Categories", "Documents" })
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"PRAGMA table_info('{table}')";
            var hasColumn = false;
            using (var reader = await cmd.ExecuteReaderAsync())
            {
                while (await reader.ReadAsync())
                {
                    if (reader.GetString(1) == "IsFavorited") { hasColumn = true; break; }
                }
            }
            if (!hasColumn)
            {
                using var alter = conn.CreateCommand();
                alter.CommandText = $"ALTER TABLE {table} ADD COLUMN IsFavorited INTEGER NOT NULL DEFAULT 0";
                await alter.ExecuteNonQueryAsync();
            }
        }

        // Add Color and IsPinned columns if missing
        foreach (var table in new[] { "Categories", "DocumentCategories", "WebDocCategories" })
        {
            using var pragmaCmd = conn.CreateCommand();
            pragmaCmd.CommandText = $"PRAGMA table_info('{table}')";
            var columns = new HashSet<string>();
            using (var reader = await pragmaCmd.ExecuteReaderAsync())
            {
                while (await reader.ReadAsync())
                    columns.Add(reader.GetString(1));
            }
            if (columns.Count > 0) // table exists
            {
                if (!columns.Contains("Color"))
                {
                    using var alter = conn.CreateCommand();
                    alter.CommandText = $"ALTER TABLE {table} ADD COLUMN Color TEXT";
                    await alter.ExecuteNonQueryAsync();
                }
                if (!columns.Contains("IsPinned"))
                {
                    using var alter = conn.CreateCommand();
                    alter.CommandText = $"ALTER TABLE {table} ADD COLUMN IsPinned INTEGER NOT NULL DEFAULT 0";
                    await alter.ExecuteNonQueryAsync();
                }
            }
        }

        // Create DocumentCategories and OfficeDocuments tables if missing
        using var checkCmd = conn.CreateCommand();
        checkCmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='DocumentCategories'";
        var exists = await checkCmd.ExecuteScalarAsync();
        if (exists is null)
        {
            using var create = conn.CreateCommand();
            create.CommandText = """
                CREATE TABLE DocumentCategories (
                    Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL DEFAULT '',
                    SortOrder INTEGER NOT NULL DEFAULT 0,
                    IsFavorited INTEGER NOT NULL DEFAULT 0,
                    IsPinned INTEGER NOT NULL DEFAULT 0,
                    Color TEXT,
                    ParentId INTEGER,
                    FOREIGN KEY (ParentId) REFERENCES DocumentCategories(Id) ON DELETE RESTRICT
                );
                CREATE UNIQUE INDEX IX_DocumentCategories_ParentId_Name ON DocumentCategories(ParentId, Name);

                CREATE TABLE OfficeDocuments (
                    Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    Title TEXT NOT NULL DEFAULT '',
                    FileName TEXT NOT NULL DEFAULT '',
                    StoredFileName TEXT NOT NULL DEFAULT '',
                    ContentType TEXT NOT NULL DEFAULT '',
                    FileSize INTEGER NOT NULL DEFAULT 0,
                    IsFavorited INTEGER NOT NULL DEFAULT 0,
                    CategoryId INTEGER NOT NULL,
                    SortOrder INTEGER NOT NULL DEFAULT 0,
                    CreatedAt TEXT NOT NULL DEFAULT '0001-01-01 00:00:00',
                    UpdatedAt TEXT NOT NULL DEFAULT '0001-01-01 00:00:00',
                    FOREIGN KEY (CategoryId) REFERENCES DocumentCategories(Id) ON DELETE CASCADE
                );
                CREATE UNIQUE INDEX IX_OfficeDocuments_CategoryId_Title ON OfficeDocuments(CategoryId, Title);
                """;
            await create.ExecuteNonQueryAsync();
        }
        // Create WebDocCategories and WebDocuments tables if missing
        using var checkWebDoc = conn.CreateCommand();
        checkWebDoc.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='WebDocCategories'";
        var webDocExists = await checkWebDoc.ExecuteScalarAsync();
        if (webDocExists is null)
        {
            using var create = conn.CreateCommand();
            create.CommandText = """
                CREATE TABLE WebDocCategories (
                    Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL DEFAULT '',
                    SortOrder INTEGER NOT NULL DEFAULT 0,
                    IsFavorited INTEGER NOT NULL DEFAULT 0,
                    IsPinned INTEGER NOT NULL DEFAULT 0,
                    Color TEXT,
                    ParentId INTEGER,
                    FOREIGN KEY (ParentId) REFERENCES WebDocCategories(Id) ON DELETE RESTRICT
                );
                CREATE UNIQUE INDEX IX_WebDocCategories_ParentId_Name ON WebDocCategories(ParentId, Name);

                CREATE TABLE WebDocuments (
                    Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    Title TEXT NOT NULL DEFAULT '',
                    HtmlContent TEXT NOT NULL DEFAULT '',
                    CategoryId INTEGER NOT NULL,
                    SortOrder INTEGER NOT NULL DEFAULT 0,
                    IsFavorited INTEGER NOT NULL DEFAULT 0,
                    CreatedAt TEXT NOT NULL DEFAULT '0001-01-01 00:00:00',
                    UpdatedAt TEXT NOT NULL DEFAULT '0001-01-01 00:00:00',
                    FOREIGN KEY (CategoryId) REFERENCES WebDocCategories(Id) ON DELETE CASCADE
                );
                CREATE UNIQUE INDEX IX_WebDocuments_CategoryId_Title ON WebDocuments(CategoryId, Title);
                """;
            await create.ExecuteNonQueryAsync();
        }

        // Create SopCategories and SopFiles tables if missing
        using var checkSopCat = conn.CreateCommand();
        checkSopCat.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='SopCategories'";
        var sopCatExists = await checkSopCat.ExecuteScalarAsync();
        if (sopCatExists is null)
        {
            using var create = conn.CreateCommand();
            create.CommandText = """
                CREATE TABLE SopCategories (
                    Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL DEFAULT '',
                    SortOrder INTEGER NOT NULL DEFAULT 0,
                    IsFavorited INTEGER NOT NULL DEFAULT 0,
                    IsPinned INTEGER NOT NULL DEFAULT 0,
                    Color TEXT,
                    ParentId INTEGER,
                    FOREIGN KEY (ParentId) REFERENCES SopCategories(Id) ON DELETE RESTRICT
                );
                CREATE UNIQUE INDEX IX_SopCategories_ParentId_Name ON SopCategories(ParentId, Name);

                CREATE TABLE SopFiles (
                    Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    Title TEXT NOT NULL DEFAULT '',
                    FileName TEXT NOT NULL DEFAULT '',
                    StoredFileName TEXT NOT NULL DEFAULT '',
                    ContentType TEXT NOT NULL DEFAULT '',
                    FileSize INTEGER NOT NULL DEFAULT 0,
                    IsFavorited INTEGER NOT NULL DEFAULT 0,
                    CategoryId INTEGER NOT NULL,
                    SortOrder INTEGER NOT NULL DEFAULT 0,
                    CreatedAt TEXT NOT NULL DEFAULT '0001-01-01 00:00:00',
                    UpdatedAt TEXT NOT NULL DEFAULT '0001-01-01 00:00:00',
                    FOREIGN KEY (CategoryId) REFERENCES SopCategories(Id) ON DELETE CASCADE
                );
                CREATE UNIQUE INDEX IX_SopFiles_CategoryId_Title ON SopFiles(CategoryId, Title);
                """;
            await create.ExecuteNonQueryAsync();
        }

        // Seed WebDocCategories from DocumentCategories (full tree)
        // Re-seeds if category count doesn't match (handles partial initial seed)
        {
            using var countWeb = conn.CreateCommand();
            countWeb.CommandText = "SELECT COUNT(*) FROM WebDocCategories";
            var webCount = Convert.ToInt32(await countWeb.ExecuteScalarAsync());

            using var countDoc = conn.CreateCommand();
            countDoc.CommandText = "SELECT COUNT(*) FROM DocumentCategories";
            var docCount = Convert.ToInt32(await countDoc.ExecuteScalarAsync());

            if (webCount != docCount && docCount > 0)
            {
                // Clear partial seed (safe because WebDocuments cascade-deletes)
                using var clear = conn.CreateCommand();
                clear.CommandText = "DELETE FROM WebDocuments; DELETE FROM WebDocCategories;";
                await clear.ExecuteNonQueryAsync();

                // Read all DocumentCategories
                var docCats = new List<(int Id, string Name, int SortOrder, int? ParentId)>();
                using var readCmd = conn.CreateCommand();
                readCmd.CommandText = "SELECT Id, Name, SortOrder, ParentId FROM DocumentCategories";
                using (var reader = await readCmd.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        docCats.Add((
                            reader.GetInt32(0),
                            reader.GetString(1),
                            reader.GetInt32(2),
                            reader.IsDBNull(3) ? null : reader.GetInt32(3)
                        ));
                    }
                }

                // Insert level by level: roots first, then children, mapping old IDs to new IDs
                var idMap = new Dictionary<int, int>(); // oldDocCatId → newWebDocCatId
                var remaining = new List<(int Id, string Name, int SortOrder, int? ParentId)>(docCats);

                while (remaining.Count > 0)
                {
                    var batch = remaining
                        .Where(c => c.ParentId is null || idMap.ContainsKey(c.ParentId.Value))
                        .ToList();

                    if (batch.Count == 0) break; // avoid infinite loop on orphans

                    foreach (var cat in batch)
                    {
                        int? newParentId = cat.ParentId.HasValue ? idMap[cat.ParentId.Value] : null;
                        using var ins = conn.CreateCommand();
                        ins.CommandText = "INSERT INTO WebDocCategories (Name, SortOrder, IsFavorited, IsPinned, Color, ParentId) VALUES (@n, @s, 0, 0, NULL, @p); SELECT last_insert_rowid();";
                        ins.Parameters.Add(new Microsoft.Data.Sqlite.SqliteParameter("@n", cat.Name));
                        ins.Parameters.Add(new Microsoft.Data.Sqlite.SqliteParameter("@s", cat.SortOrder));
                        ins.Parameters.Add(new Microsoft.Data.Sqlite.SqliteParameter("@p", (object?)newParentId ?? DBNull.Value));
                        var newId = Convert.ToInt32(await ins.ExecuteScalarAsync());
                        idMap[cat.Id] = newId;
                    }

                    foreach (var cat in batch)
                        remaining.Remove(cat);
                }
            }
        }
    }
      catch (Exception ex)
      {
          var logger = app.Services.GetRequiredService<ILoggerFactory>().CreateLogger("DecoSOP.Migration");
          logger.LogError(ex, "Database schema migration failed. The app will continue but some features may not work correctly until the database is updated.");
      }
    }
    finally
    {
        await conn.CloseAsync();
    }

    // Seed demo data if requested via --seed-demo flag
    if (args.Contains("--seed-demo"))
        await DemoDataService.SeedDemoDataAsync(db);

    // Import files from directory if requested via CLI flags
    var importSopsIdx = Array.IndexOf(args, "--import-sops");
    if (importSopsIdx >= 0 && importSopsIdx + 1 < args.Length)
    {
        var sopSourceDir = args[importSopsIdx + 1];
        var sopUploadsDir = Path.Combine(dataDir, "sop-uploads");
        FileImportService.ImportSopFiles(dbPath, sopSourceDir, sopUploadsDir);
    }

    var importDocsIdx = Array.IndexOf(args, "--import-docs");
    if (importDocsIdx >= 0 && importDocsIdx + 1 < args.Length)
    {
        var docSourceDir = args[importDocsIdx + 1];
        var docUploadsDir = Path.Combine(dataDir, "uploads");
        FileImportService.ImportDocumentFiles(dbPath, docSourceDir, docUploadsDir);
    }

    // Exit after CLI operations (don't start the web server)
    if (args.Contains("--seed-demo") || importSopsIdx >= 0 || importDocsIdx >= 0)
        return;

    // Ensure uploads directories exist
    DocumentService.GetUploadDirectory();
    SopFileService.GetUploadDirectory();
}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

// File download endpoint for office documents
app.MapGet("/api/documents/{id:int}/download", async (int id, AppDbContext db) =>
{
    var doc = await db.OfficeDocuments.FindAsync(id);
    if (doc is null) return Results.NotFound();

    var filePath = Path.Combine(DocumentService.GetUploadDirectory(), doc.StoredFileName);
    if (!File.Exists(filePath)) return Results.NotFound();

    return Results.File(filePath, doc.ContentType, doc.FileName);
});

// File preview endpoint — serves inline (no Content-Disposition: attachment)
app.MapGet("/api/documents/{id:int}/preview", async (int id, AppDbContext db) =>
{
    var doc = await db.OfficeDocuments.FindAsync(id);
    if (doc is null) return Results.NotFound();

    var filePath = Path.Combine(DocumentService.GetUploadDirectory(), doc.StoredFileName);
    if (!File.Exists(filePath)) return Results.NotFound();

    return Results.File(filePath, doc.ContentType, enableRangeProcessing: true);
});

// HTML preview for Office documents (DOCX, XLSX, DOC → converted to HTML)
app.MapGet("/api/documents/{id:int}/preview-html", async (int id, AppDbContext db) =>
{
    var doc = await db.OfficeDocuments.FindAsync(id);
    if (doc is null) return Results.NotFound();

    var filePath = Path.Combine(DocumentService.GetUploadDirectory(), doc.StoredFileName);
    if (!File.Exists(filePath)) return Results.NotFound();

    var html = DocumentPreviewService.GenerateHtmlPreview(filePath, doc.Title);
    return Results.Content(html, "text/html");
});

// PDF preview for Office documents — converts via LibreOffice on first access, then caches
app.MapGet("/api/documents/{id:int}/preview-pdf", async (int id, AppDbContext db) =>
{
    var doc = await db.OfficeDocuments.FindAsync(id);
    if (doc is null) return Results.NotFound();

    var filePath = Path.Combine(DocumentService.GetUploadDirectory(), doc.StoredFileName);
    if (!File.Exists(filePath)) return Results.NotFound();

    var pdfPath = await PdfConversionService.GetOrCreatePdfAsync(filePath, doc.StoredFileName);
    if (pdfPath is null)
        return Results.Problem("PDF conversion failed. Ensure LibreOffice is installed.");

    return Results.File(pdfPath, "application/pdf", enableRangeProcessing: true);
});

// File download endpoint for SOP files
app.MapGet("/api/sops/{id:int}/download", async (int id, AppDbContext db) =>
{
    var doc = await db.SopFiles.FindAsync(id);
    if (doc is null) return Results.NotFound();

    var filePath = Path.Combine(SopFileService.GetUploadDirectory(), doc.StoredFileName);
    if (!File.Exists(filePath)) return Results.NotFound();

    return Results.File(filePath, doc.ContentType, doc.FileName);
});

app.MapGet("/api/sops/{id:int}/preview", async (int id, AppDbContext db) =>
{
    var doc = await db.SopFiles.FindAsync(id);
    if (doc is null) return Results.NotFound();

    var filePath = Path.Combine(SopFileService.GetUploadDirectory(), doc.StoredFileName);
    if (!File.Exists(filePath)) return Results.NotFound();

    return Results.File(filePath, doc.ContentType, enableRangeProcessing: true);
});

app.MapGet("/api/sops/{id:int}/preview-html", async (int id, AppDbContext db) =>
{
    var doc = await db.SopFiles.FindAsync(id);
    if (doc is null) return Results.NotFound();

    var filePath = Path.Combine(SopFileService.GetUploadDirectory(), doc.StoredFileName);
    if (!File.Exists(filePath)) return Results.NotFound();

    var html = DocumentPreviewService.GenerateHtmlPreview(filePath, doc.Title);
    return Results.Content(html, "text/html");
});

app.MapGet("/api/sops/{id:int}/preview-pdf", async (int id, AppDbContext db) =>
{
    var doc = await db.SopFiles.FindAsync(id);
    if (doc is null) return Results.NotFound();

    var filePath = Path.Combine(SopFileService.GetUploadDirectory(), doc.StoredFileName);
    if (!File.Exists(filePath)) return Results.NotFound();

    var pdfPath = await PdfConversionService.GetOrCreatePdfAsync(filePath, doc.StoredFileName);
    if (pdfPath is null)
        return Results.Problem("PDF conversion failed. Ensure LibreOffice is installed.");

    return Results.File(pdfPath, "application/pdf", enableRangeProcessing: true);
});

// Database export endpoint
app.MapGet("/api/settings/export-db", () =>
{
    if (!File.Exists(dbPath)) return Results.NotFound();
    return Results.File(dbPath, "application/octet-stream", "decosop-backup.db");
});

app.Run();
