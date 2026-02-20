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
            categoryColumns: "Name, SortOrder, IsFavorited, IsPinned, Color, ParentId",
            categoryValues: "@name, @sort, 0, 0, NULL, @parent",
            fileColumns: "Title, FileName, StoredFileName, ContentType, FileSize, IsFavorited, CategoryId, SortOrder, CreatedAt, UpdatedAt");
    }

    /// <summary>
    /// Import files from sourceDir into the database as Documents.
    /// </summary>
    public static int ImportDocumentFiles(string dbPath, string sourceDir, string uploadsDir)
    {
        return ImportFiles(dbPath, sourceDir, uploadsDir,
            categoryTable: "DocumentCategories",
            fileTable: "OfficeDocuments",
            categoryColumns: "Name, SortOrder, IsFavorited, IsPinned, Color, ParentId",
            categoryValues: "@name, @sort, 0, 0, NULL, @parent",
            fileColumns: "Title, FileName, StoredFileName, ContentType, FileSize, IsFavorited, CategoryId, SortOrder, CreatedAt, UpdatedAt");
    }

    private static int ImportFiles(
        string dbPath, string sourceDir, string uploadsDir,
        string categoryTable, string fileTable,
        string categoryColumns, string categoryValues,
        string fileColumns)
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

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        var catSortCounters = new Dictionary<int?, int>();
        var docSortCounters = new Dictionary<int, int>();
        var usedTitles = new Dictionary<int, HashSet<string>>();
        var now = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");
        int imported = 0;

        foreach (var (relPath, fullPath) in files)
        {
            var catId = GetCategoryForPath(conn, relPath, catSortCounters, categoryTable, categoryColumns, categoryValues);
            var originalFilename = Path.GetFileName(fullPath);
            var title = CleanTitle(originalFilename);
            var contentType = GetContentType(originalFilename);
            var fileSize = new FileInfo(fullPath).Length;

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
                // Insert DB record
                using var ins = conn.CreateCommand();
                ins.CommandText = $"INSERT INTO {fileTable} ({fileColumns}) VALUES (@title, @fileName, '', @contentType, @fileSize, 0, @catId, @sortOrder, @now, @now); SELECT last_insert_rowid();";
                ins.Parameters.AddWithValue("@title", title);
                ins.Parameters.AddWithValue("@fileName", originalFilename);
                ins.Parameters.AddWithValue("@contentType", contentType);
                ins.Parameters.AddWithValue("@fileSize", fileSize);
                ins.Parameters.AddWithValue("@catId", catId);
                ins.Parameters.AddWithValue("@sortOrder", sortOrder);
                ins.Parameters.AddWithValue("@now", now);
                var docId = Convert.ToInt32(ins.ExecuteScalar());

                // Copy file with id-prefixed name
                var safeName = Regex.Replace(originalFilename, @"[<>:""/\\|?*]", "_");
                var storedName = $"{docId}_{safeName}";
                var destPath = Path.Combine(uploadsDir, storedName);
                File.Copy(fullPath, destPath, overwrite: true);

                // Update stored filename
                using var upd = conn.CreateCommand();
                upd.CommandText = $"UPDATE {fileTable} SET StoredFileName = @stored WHERE Id = @id";
                upd.Parameters.AddWithValue("@stored", storedName);
                upd.Parameters.AddWithValue("@id", docId);
                upd.ExecuteNonQuery();

                imported++;

                if (imported % 100 == 0)
                    Console.WriteLine($"  Imported {imported}/{files.Count}...");
            }
            catch (SqliteException ex) when (ex.SqliteErrorCode == 19)
            {
                Console.WriteLine($"  DUPLICATE: {title} in category {catId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ERROR: {fullPath}: {ex.Message}");
            }
        }

        Console.WriteLine($"Import complete: {imported} files imported into {categoryTable}/{fileTable}");
        return imported;
    }

    private static int GetCategoryForPath(
        SqliteConnection conn, string relPath,
        Dictionary<int?, int> sortCounters,
        string table, string columns, string values)
    {
        var parts = relPath.Replace('\\', '/').Split('/');

        if (parts.Length <= 1)
            return EnsureCategory(conn, "General", null, sortCounters, table, columns, values);

        var dirParts = parts[..^1]; // all except filename
        int? parentId = null;
        for (var i = 0; i < dirParts.Length; i++)
        {
            var cleanName = CleanDirName(dirParts[i]);
            parentId = EnsureCategory(conn, cleanName, parentId, sortCounters, table, columns, values);
        }

        return parentId!.Value;
    }

    private static int EnsureCategory(
        SqliteConnection conn, string name, int? parentId,
        Dictionary<int?, int> sortCounters,
        string table, string columns, string values)
    {
        // Check if exists
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
            return Convert.ToInt32(existing);

        // Create new category
        var sortOrder = sortCounters.GetValueOrDefault(parentId, 0);
        sortCounters[parentId] = sortOrder + 1;

        using var ins = conn.CreateCommand();
        ins.CommandText = $"INSERT INTO {table} ({columns}) VALUES ({values}); SELECT last_insert_rowid();";
        ins.Parameters.AddWithValue("@name", name);
        ins.Parameters.AddWithValue("@sort", sortOrder);
        ins.Parameters.AddWithValue("@parent", parentId.HasValue ? parentId.Value : DBNull.Value);
        return Convert.ToInt32(ins.ExecuteScalar());
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
