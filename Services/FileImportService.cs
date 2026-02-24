using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace DecoSOP.Services;

/// <summary>
/// Imports files from a directory tree into the database, creating categories
/// from folder structure. Used by --import-sops and --import-docs CLI flags.
/// </summary>
public static class FileImportService
{
    private static readonly string[] SkipPatterns =
    [
        @"^zz\s*archive$",
        @"^z\s*archive$",
        @"^x\s*old$",
        @"^z\s*old\b",
        @"^poss\s*older\b",
        @"^!+\s*.*to\s+go\s+thru",
        @"to\s+go\s+thru",
        @"^archive$"
    ];

    private static readonly HashSet<string> IncludeExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
        ".txt", ".csv", ".png", ".jpg", ".jpeg", ".gif", ".zip",
        ".rtf", ".odt", ".ods"
    };

    /// <summary>
    /// Import files from sourceDir into the database as SOP files.
    /// </summary>
    public static int ImportSopFiles(string dbPath, string sourceDir, string uploadsDir)
    {
        return ImportFiles(dbPath, sourceDir, uploadsDir,
            categoryTable: "SopCategories",
            fileTable: "SopFiles",
            categoryColumns: "Name, SortOrder, ParentId",
            categoryValues: "@name, @sort, @parent",
            fileColumns: "Title, FileName, StoredFileName, ContentType, FileSize, CategoryId, SortOrder, CreatedAt, UpdatedAt",
            webCategoryTable: "Categories",
            webDocTable: "Documents");
    }

    /// <summary>
    /// Import files from sourceDir into the database as Documents.
    /// </summary>
    public static int ImportDocumentFiles(string dbPath, string sourceDir, string uploadsDir)
    {
        return ImportFiles(dbPath, sourceDir, uploadsDir,
            categoryTable: "DocumentCategories",
            fileTable: "OfficeDocuments",
            categoryColumns: "Name, SortOrder, ParentId",
            categoryValues: "@name, @sort, @parent",
            fileColumns: "Title, FileName, StoredFileName, ContentType, FileSize, CategoryId, SortOrder, CreatedAt, UpdatedAt",
            webCategoryTable: "WebDocCategories",
            webDocTable: "WebDocuments");
    }

    private static int ImportFiles(
        string dbPath, string sourceDir, string uploadsDir,
        string categoryTable, string fileTable,
        string categoryColumns, string categoryValues,
        string fileColumns,
        string webCategoryTable, string webDocTable)
    {
        if (!Directory.Exists(sourceDir))
        {
            Console.WriteLine($"ERROR: Source directory not found: {sourceDir}");
            return 0;
        }

        Directory.CreateDirectory(uploadsDir);

        var files = WalkFiles(sourceDir).ToList();
        Console.WriteLine($"Found {files.Count} files to import from {sourceDir}");

        if (files.Count == 0)
        {
            Console.WriteLine("No matching files found.");
            return 0;
        }

        // Log file for debugging (process runs hidden during install)
        var logFile = Path.Combine(Path.GetDirectoryName(dbPath)!, "import-log.txt");
        void Log(string msg)
        {
            var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
            Console.WriteLine(line);
            try { File.AppendAllText(logFile, line + Environment.NewLine); } catch { }
        }

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        // Optimize SQLite for bulk import
        // NOTE: Use raw SQL for transactions (BEGIN/COMMIT) instead of conn.BeginTransaction()
        // because Microsoft.Data.Sqlite requires every command to have cmd.Transaction set
        // when using the .NET transaction API, which breaks helper methods.
        using (var pragma = conn.CreateCommand())
        {
            pragma.CommandText = "PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL; PRAGMA temp_store=MEMORY; PRAGMA cache_size=-64000;";
            pragma.ExecuteNonQuery();
        }

        // Clear stale data from previous imports to prevent ghost records
        Log($"Clearing existing data from {fileTable}, {categoryTable}, {webDocTable}, {webCategoryTable}...");
        using (var cleanup = conn.CreateCommand())
        {
            cleanup.CommandText = $"""
                DELETE FROM {webDocTable};
                DELETE FROM {webCategoryTable};
                DELETE FROM {fileTable};
                DELETE FROM {categoryTable};
                """;
            cleanup.ExecuteNonQuery();
        }
        Log("Stale data cleared.");

        var catSortCounters = new Dictionary<int, int>();
        var catCache = new Dictionary<(string Name, int ParentKey), int>();
        var docSortCounters = new Dictionary<int, int>();
        var usedTitles = new Dictionary<int, HashSet<string>>();
        var now = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");
        int imported = 0;
        int errors = 0;

        // Web content generation tracking
        var webCatMap = new Dictionary<int, int>();  // file category ID → web category ID
        var webCatSortCounters = new Dictionary<int, int>();
        var webDocSortCounters = new Dictionary<int, int>();
        int webGenerated = 0;

        // Status file for installer progress display
        var statusFile = Path.Combine(Path.GetDirectoryName(dbPath)!, "import-status.txt");

        // Process in transaction batches for performance (raw SQL transactions)
        const int batchSize = 200;
        ExecSql(conn, "BEGIN TRANSACTION");

        foreach (var (relPath, fullPath) in files)
        {
            var catId = GetCategoryForPath(conn, relPath, catSortCounters, catCache, categoryTable, categoryColumns, categoryValues);
            var originalFilename = Path.GetFileName(fullPath);
            var title = CleanTitle(originalFilename);
            var contentType = GetContentType(originalFilename);
            var fileSize = new FileInfo(fullPath).Length;

            // Log first few files for debugging
            if (imported + errors < 5)
                Log($"  File: relPath={relPath}, catId={catId}, title={title}");

            // Ensure unique title within category
            if (!usedTitles.ContainsKey(catId))
                usedTitles[catId] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var baseTitle = title;
            var counter = 1;
            while (usedTitles[catId].Contains(title))
            {
                counter++;
                title = $"{baseTitle} ({counter})";
            }
            usedTitles[catId].Add(title);

            var sortOrder = docSortCounters.GetValueOrDefault(catId, 0);
            docSortCounters[catId] = sortOrder + 1;

            try
            {
                // Copy file preserving original directory structure
                // Skip copy when source dir IS the uploads dir (files already in place)
                var storedName = relPath.Replace('\\', '/');
                var destPath = Path.Combine(uploadsDir, relPath);
                if (!string.Equals(Path.GetFullPath(fullPath), Path.GetFullPath(destPath), StringComparison.OrdinalIgnoreCase))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(destPath)!);
                    File.Copy(fullPath, destPath, overwrite: true);
                }

                // Insert DB record with StoredFileName already set
                using var ins = conn.CreateCommand();
                ins.CommandText = $"INSERT INTO {fileTable} ({fileColumns}) VALUES (@title, @fileName, @storedName, @contentType, @fileSize, @catId, @sortOrder, @now, @now); SELECT last_insert_rowid();";
                ins.Parameters.AddWithValue("@title", title);
                ins.Parameters.AddWithValue("@fileName", originalFilename);
                ins.Parameters.AddWithValue("@storedName", storedName);
                ins.Parameters.AddWithValue("@contentType", contentType);
                ins.Parameters.AddWithValue("@fileSize", fileSize);
                ins.Parameters.AddWithValue("@catId", catId);
                ins.Parameters.AddWithValue("@sortOrder", sortOrder);
                ins.Parameters.AddWithValue("@now", now);
                var docId = Convert.ToInt32(ins.ExecuteScalar());

                imported++;

                // Write progress for installer status display
                try { File.WriteAllText(statusFile, $"{imported}/{files.Count} — {title}"); } catch { }

                // Auto-generate web content
                try
                {
                    var webCatId = GetWebCategoryForFileCat(
                        conn, catId, webCatMap, webCatSortCounters,
                        categoryTable, webCategoryTable);

                    var htmlContent = WebContentExtractor.ExtractHtml(fullPath, title);

                    var webSortOrder = webDocSortCounters.GetValueOrDefault(webCatId, 0);
                    webDocSortCounters[webCatId] = webSortOrder + 1;

                    using var webIns = conn.CreateCommand();
                    webIns.CommandText = $"INSERT OR IGNORE INTO {webDocTable} (Title, HtmlContent, CategoryId, SortOrder, CreatedAt, UpdatedAt) VALUES (@title, @html, @catId, @sort, @now, @now)";
                    webIns.Parameters.AddWithValue("@title", title);
                    webIns.Parameters.AddWithValue("@html", htmlContent);
                    webIns.Parameters.AddWithValue("@catId", webCatId);
                    webIns.Parameters.AddWithValue("@sort", webSortOrder);
                    webIns.Parameters.AddWithValue("@now", now);
                    if (webIns.ExecuteNonQuery() > 0)
                        webGenerated++;
                }
                catch (Exception webEx)
                {
                    Log($"  WARNING: Web content generation failed for {title}: {webEx.Message}");
                }

                // Commit in batches for performance
                if (imported % batchSize == 0)
                {
                    ExecSql(conn, "COMMIT");
                    ExecSql(conn, "BEGIN TRANSACTION");
                    Log($"  Imported {imported}/{files.Count}...");
                }
            }
            catch (SqliteException ex) when (ex.SqliteErrorCode == 19)
            {
                Log($"  DUPLICATE: {title} in category {catId}");
                errors++;
            }
            catch (Exception ex)
            {
                Log($"  ERROR: {fullPath}: {ex.Message}");
                errors++;
            }
        }

        // Commit remaining files
        ExecSql(conn, "COMMIT");

        Log($"Import complete: {imported} files imported, {errors} errors, into {categoryTable}/{fileTable}");
        Log($"  Web content generated: {webGenerated} into {webCategoryTable}/{webDocTable}");

        // Signal completion to installer
        try { File.WriteAllText(statusFile, "COMPLETE"); } catch { }

        return imported;
    }

    private static int GetCategoryForPath(
        SqliteConnection conn, string relPath,
        Dictionary<int, int> sortCounters,
        Dictionary<(string Name, int ParentKey), int> cache,
        string table, string columns, string values)
    {
        var parts = relPath.Replace('\\', '/').Split('/');

        if (parts.Length <= 1)
            return EnsureCategory(conn, "General", null, sortCounters, cache, table, columns, values);

        var dirParts = parts[..^1]; // all except filename
        int? parentId = null;
        for (var i = 0; i < dirParts.Length; i++)
        {
            var cleanName = CleanDirName(dirParts[i]);
            parentId = EnsureCategory(conn, cleanName, parentId, sortCounters, cache, table, columns, values);
        }

        return parentId!.Value;
    }

    private static int EnsureCategory(
        SqliteConnection conn, string name, int? parentId,
        Dictionary<int, int> sortCounters,
        Dictionary<(string Name, int ParentKey), int> cache,
        string table, string columns, string values)
    {
        // Check in-memory cache first (avoids DB round-trip for repeated lookups)
        var cacheKey = (name, parentId ?? -1);
        if (cache.TryGetValue(cacheKey, out var cachedId))
            return cachedId;

        // Check if exists in DB
        using var check = conn.CreateCommand();
        if (parentId is null)
        {
            check.CommandText = $"SELECT Id FROM {table} WHERE Name = @name AND ParentId IS NULL";
            check.Parameters.AddWithValue("@name", name);
        }
        else
        {
            check.CommandText = $"SELECT Id FROM {table} WHERE Name = @name AND ParentId = @parent";
            check.Parameters.AddWithValue("@name", name);
            check.Parameters.AddWithValue("@parent", parentId.Value);
        }

        var existing = check.ExecuteScalar();
        if (existing is not null)
        {
            var id = Convert.ToInt32(existing);
            cache[cacheKey] = id;
            return id;
        }

        // Create new category (use -1 as dictionary key for null parentId since Dictionary doesn't allow null keys)
        var sortKey = parentId ?? -1;
        var sortOrder = sortCounters.GetValueOrDefault(sortKey, 0);
        sortCounters[sortKey] = sortOrder + 1;

        using var ins = conn.CreateCommand();
        ins.CommandText = $"INSERT INTO {table} ({columns}) VALUES ({values}); SELECT last_insert_rowid();";
        ins.Parameters.AddWithValue("@name", name);
        ins.Parameters.AddWithValue("@sort", sortOrder);
        ins.Parameters.AddWithValue("@parent", parentId.HasValue ? parentId.Value : DBNull.Value);
        var newId = Convert.ToInt32(ins.ExecuteScalar());
        cache[cacheKey] = newId;
        return newId;
    }

    /// <summary>
    /// Given a file-category ID, ensure a matching category exists in the web-category table.
    /// Walks the ancestor chain and mirrors each level. Results are cached in webCatMap.
    /// </summary>
    private static int GetWebCategoryForFileCat(
        SqliteConnection conn, int fileCatId,
        Dictionary<int, int> webCatMap,
        Dictionary<int, int> webCatSortCounters,
        string fileCatTable, string webCatTable)
    {
        if (webCatMap.TryGetValue(fileCatId, out var cached))
            return cached;

        // Read the file category's full ancestor chain
        var chain = new List<(int Id, string Name, int? ParentId)>();
        var currentId = fileCatId;
        while (true)
        {
            using var read = conn.CreateCommand();
            read.CommandText = $"SELECT Id, Name, ParentId FROM {fileCatTable} WHERE Id = @id";
            read.Parameters.AddWithValue("@id", currentId);
            using var reader = read.ExecuteReader();
            if (!reader.Read()) break;
            var id = reader.GetInt32(0);
            var name = reader.GetString(1);
            var parentId = reader.IsDBNull(2) ? (int?)null : reader.GetInt32(2);
            chain.Insert(0, (id, name, parentId));
            if (parentId is null) break;
            currentId = parentId.Value;
        }

        // Ensure each level exists in the web category table
        int? webParentId = null;
        foreach (var (id, name, _) in chain)
        {
            if (webCatMap.TryGetValue(id, out var existing))
            {
                webParentId = existing;
                continue;
            }

            // Check if web category already exists
            using var check = conn.CreateCommand();
            if (webParentId is null)
            {
                check.CommandText = $"SELECT Id FROM {webCatTable} WHERE Name = @name AND ParentId IS NULL";
                check.Parameters.AddWithValue("@name", name);
            }
            else
            {
                check.CommandText = $"SELECT Id FROM {webCatTable} WHERE Name = @name AND ParentId = @parent";
                check.Parameters.AddWithValue("@name", name);
                check.Parameters.AddWithValue("@parent", webParentId.Value);
            }
            var existingWebId = check.ExecuteScalar();

            if (existingWebId is not null)
            {
                var webId = Convert.ToInt32(existingWebId);
                webCatMap[id] = webId;
                webParentId = webId;
            }
            else
            {
                var sortKey = webParentId ?? -1;
                var sortOrder = webCatSortCounters.GetValueOrDefault(sortKey, 0);
                webCatSortCounters[sortKey] = sortOrder + 1;

                using var ins = conn.CreateCommand();
                ins.CommandText = $"INSERT INTO {webCatTable} (Name, SortOrder, ParentId) VALUES (@name, @sort, @parent); SELECT last_insert_rowid();";
                ins.Parameters.AddWithValue("@name", name);
                ins.Parameters.AddWithValue("@sort", sortOrder);
                ins.Parameters.AddWithValue("@parent", webParentId.HasValue ? webParentId.Value : DBNull.Value);
                var newWebId = Convert.ToInt32(ins.ExecuteScalar());
                webCatMap[id] = newWebId;
                webParentId = newWebId;
            }
        }

        return webParentId!.Value;
    }

    /// <summary>Execute a raw SQL statement (used for BEGIN/COMMIT to avoid .NET transaction API issues).</summary>
    private static void ExecSql(SqliteConnection conn, string sql)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();
    }

    private static IEnumerable<(string RelPath, string FullPath)> WalkFiles(string baseDir)
    {
        var baseInfo = new DirectoryInfo(baseDir);
        return WalkDirectory(baseInfo, baseInfo.FullName);
    }

    private static IEnumerable<(string RelPath, string FullPath)> WalkDirectory(DirectoryInfo dir, string basePath)
    {
        FileInfo[] files;
        try { files = dir.GetFiles(); }
        catch (UnauthorizedAccessException) { yield break; }

        foreach (var file in files.OrderBy(f => f.Name))
        {
            if (file.Name.StartsWith("~$"))
                continue;

            if (IncludeExtensions.Count > 0 && !IncludeExtensions.Contains(file.Extension))
                continue;

            var relPath = Path.GetRelativePath(basePath, file.FullName);
            yield return (relPath, file.FullName);
        }

        DirectoryInfo[] subdirs;
        try { subdirs = dir.GetDirectories(); }
        catch (UnauthorizedAccessException) { yield break; }

        foreach (var subdir in subdirs.OrderBy(d => d.Name))
        {
            if (ShouldSkipDir(subdir.Name))
                continue;

            foreach (var item in WalkDirectory(subdir, basePath))
                yield return item;
        }
    }

    private static bool ShouldSkipDir(string dirname)
    {
        foreach (var pattern in SkipPatterns)
        {
            if (Regex.IsMatch(dirname, pattern, RegexOptions.IgnoreCase))
                return true;
        }
        return false;
    }

    private static string CleanDirName(string dirname)
    {
        var name = dirname;
        name = Regex.Replace(name, @"^[~!^@]+\s*", "");
        name = Regex.Replace(name, @"^\d+\s+", "");
        name = Regex.Replace(name, @"\s+", " ").Trim();
        return string.IsNullOrEmpty(name) ? dirname : name;
    }

    private static string CleanTitle(string filename)
    {
        var name = Path.GetFileNameWithoutExtension(filename);
        name = Regex.Replace(name, @"^[~!^]+\s*", "");
        name = name.Replace('_', ' ');
        name = Regex.Replace(name, @"\s+", " ").Trim();
        return string.IsNullOrEmpty(name) ? filename : name;
    }

    private static string GetContentType(string filename)
    {
        var ext = Path.GetExtension(filename).ToLowerInvariant();
        return ext switch
        {
            ".pdf" => "application/pdf",
            ".doc" => "application/msword",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".xls" => "application/vnd.ms-excel",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".ppt" => "application/vnd.ms-powerpoint",
            ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".txt" => "text/plain",
            ".csv" => "text/csv",
            ".rtf" => "application/rtf",
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".zip" => "application/zip",
            ".odt" => "application/vnd.oasis.opendocument.text",
            ".ods" => "application/vnd.oasis.opendocument.spreadsheet",
            _ => "application/octet-stream"
        };
    }
}
