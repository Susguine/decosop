"""
Import office documents from a directory into the DecoSOP SQLite database.

Copies files into the uploads/ directory and creates metadata records in the
OfficeDocuments table. Mirrors the source folder structure as DocumentCategories.

Unlike the SOP importer, this does NOT convert files — they're stored as-is
for download. Supports any file type.
"""

import mimetypes
import os
import re
import shutil
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path

# --- Configuration ---
DOCUMENTS_DIR = r"C:\Users\JenniferD\OneDrive - ZO Organized (1)\!   Admin@deco FILES\DOCUMENTS"
DB_PATH = r"C:\Users\JenniferD\source\repos\DecoSOP\decosop.db"
UPLOADS_DIR = r"C:\Users\JenniferD\source\repos\DecoSOP\uploads"

# Directories to skip (matched as standalone directory names, case-insensitive)
SKIP_PATTERNS = [
    re.compile(r"^zz\s*archive$", re.IGNORECASE),
    re.compile(r"^z\s*archive$", re.IGNORECASE),
    re.compile(r"^x\s*old$", re.IGNORECASE),
    re.compile(r"^z\s*old\b", re.IGNORECASE),
    re.compile(r"^archive$", re.IGNORECASE),
]

# File extensions to include (empty = include everything)
# Add or remove extensions as needed
INCLUDE_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".txt", ".csv", ".png", ".jpg", ".jpeg", ".gif", ".zip",
    ".rtf", ".odt", ".ods",
}

# Files at the root level (not in any subfolder) go to this category
ROOT_CATEGORY = "General"


def should_skip_dir(dirname: str) -> bool:
    for pattern in SKIP_PATTERNS:
        if pattern.search(dirname):
            return True
    return False


def clean_dirname(dirname: str) -> str:
    name = dirname
    name = re.sub(r"^[~!^@]+\s*", "", name)
    name = re.sub(r"^\d+\s+", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else dirname


def clean_title(filename: str) -> str:
    name = Path(filename).stem
    name = re.sub(r"^[~!^]+\s*", "", name)
    name = name.replace("_", " ")
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else filename


def get_content_type(filepath: str) -> str:
    mime, _ = mimetypes.guess_type(filepath)
    return mime or "application/octet-stream"


def create_tables(cursor):
    """Create the DocumentCategories and OfficeDocuments tables if they don't exist."""
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS DocumentCategories (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Name TEXT NOT NULL DEFAULT '',
            SortOrder INTEGER NOT NULL DEFAULT 0,
            IsFavorited INTEGER NOT NULL DEFAULT 0,
            ParentId INTEGER,
            FOREIGN KEY (ParentId) REFERENCES DocumentCategories(Id) ON DELETE RESTRICT
        )
    """)
    cursor.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS IX_DocumentCategories_ParentId_Name "
        "ON DocumentCategories(ParentId, Name)"
    )

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS OfficeDocuments (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        )
    """)
    cursor.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS IX_OfficeDocuments_CategoryId_Title "
        "ON OfficeDocuments(CategoryId, Title)"
    )


def ensure_category(cursor, name: str, parent_id: int | None, sort_counters: dict) -> int:
    if parent_id is None:
        cursor.execute(
            "SELECT Id FROM DocumentCategories WHERE Name = ? AND ParentId IS NULL", (name,)
        )
    else:
        cursor.execute(
            "SELECT Id FROM DocumentCategories WHERE Name = ? AND ParentId = ?", (name, parent_id)
        )
    row = cursor.fetchone()
    if row:
        return row[0]

    sort_key = (parent_id,)
    sort_order = sort_counters.get(sort_key, 0)
    sort_counters[sort_key] = sort_order + 1

    cursor.execute(
        "INSERT INTO DocumentCategories (Name, SortOrder, IsFavorited, ParentId) VALUES (?, ?, 0, ?)",
        (name, sort_order, parent_id),
    )
    return cursor.lastrowid


def get_category_for_path(cursor, rel_path: str, sort_counters: dict) -> int:
    parts = Path(rel_path).parts

    if len(parts) <= 1:
        return ensure_category(cursor, ROOT_CATEGORY, None, sort_counters)

    dir_parts = parts[:-1]
    parent_id = None
    for dirname in dir_parts:
        clean_name = clean_dirname(dirname)
        parent_id = ensure_category(cursor, clean_name, parent_id, sort_counters)

    return parent_id


def walk_documents(base_dir: str):
    """Walk the directory, yielding (rel_path, full_path) tuples."""
    base = Path(base_dir)
    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if not should_skip_dir(d)]

        for filename in sorted(files):
            if filename.startswith("~$"):
                continue

            ext = Path(filename).suffix.lower()
            if INCLUDE_EXTENSIONS and ext not in INCLUDE_EXTENSIONS:
                continue

            full_path = os.path.join(root, filename)
            rel_path = os.path.relpath(full_path, base)
            yield rel_path, full_path


def main():
    if DOCUMENTS_DIR == r"C:\Path\To\Your\Documents":
        print("ERROR: Edit DOCUMENTS_DIR at the top of this script to point to your documents folder.")
        sys.exit(1)

    print(f"Scanning: {DOCUMENTS_DIR}")
    print(f"Database: {DB_PATH}")
    print(f"Uploads:  {UPLOADS_DIR}")
    print()

    # Collect all files
    files = list(walk_documents(DOCUMENTS_DIR))
    print(f"Found {len(files)} files to import")

    by_ext = {}
    for rel_path, _ in files:
        ext = Path(rel_path).suffix.lower()
        by_ext[ext] = by_ext.get(ext, 0) + 1
    for ext, count in sorted(by_ext.items()):
        print(f"  {ext}: {count}")
    print()

    if not files:
        print("No files found. Check DOCUMENTS_DIR and INCLUDE_EXTENSIONS.")
        sys.exit(0)

    # Ensure uploads directory exists
    os.makedirs(UPLOADS_DIR, exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    create_tables(cursor)
    conn.commit()

    cat_sort_counters = {}
    doc_sort_counters = {}
    used_titles = {}

    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    imported = 0
    skipped = 0
    errors = 0

    for idx, (rel_path, full_path) in enumerate(files):
        if (idx + 1) % 50 == 0:
            print(f"  Processing {idx + 1}/{len(files)}...")

        cat_id = get_category_for_path(cursor, rel_path, cat_sort_counters)
        title = clean_title(os.path.basename(rel_path))
        original_filename = os.path.basename(rel_path)
        content_type = get_content_type(full_path)
        file_size = os.path.getsize(full_path)

        # Ensure unique title within category
        if cat_id not in used_titles:
            used_titles[cat_id] = set()
        base_title = title
        counter = 1
        while title in used_titles[cat_id]:
            counter += 1
            title = f"{base_title} ({counter})"
        used_titles[cat_id].add(title)

        sort_order = doc_sort_counters.get(cat_id, 0)
        doc_sort_counters[cat_id] = sort_order + 1

        try:
            # Insert DB record first to get the Id
            cursor.execute(
                """INSERT INTO OfficeDocuments
                   (Title, FileName, StoredFileName, ContentType, FileSize,
                    IsFavorited, CategoryId, SortOrder, CreatedAt, UpdatedAt)
                   VALUES (?, ?, '', ?, ?, 0, ?, ?, ?, ?)""",
                (title, original_filename, content_type, file_size,
                 cat_id, sort_order, now, now),
            )
            doc_id = cursor.lastrowid

            # Build stored filename and copy file
            safe_name = re.sub(r'[<>:"/\\|?*]', '_', original_filename)
            stored_name = f"{doc_id}_{safe_name}"
            dest_path = os.path.join(UPLOADS_DIR, stored_name)
            shutil.copy2(full_path, dest_path)

            # Update the stored filename
            cursor.execute(
                "UPDATE OfficeDocuments SET StoredFileName = ? WHERE Id = ?",
                (stored_name, doc_id),
            )
            imported += 1

        except sqlite3.IntegrityError as e:
            print(f"  DUPLICATE: {title} in category {cat_id}: {e}")
            errors += 1
        except Exception as e:
            print(f"  ERROR copying {full_path}: {e}")
            errors += 1

    conn.commit()

    # Print summary
    print()
    print("=" * 60)
    print(f"Import complete!")
    print(f"  Documents imported: {imported}")
    print(f"  Skipped: {skipped}")
    print(f"  Errors: {errors}")

    cursor.execute("SELECT COUNT(*) FROM DocumentCategories")
    total_cats = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM DocumentCategories WHERE ParentId IS NULL")
    root_cats = cursor.fetchone()[0]
    print(f"  Categories: {total_cats} total ({root_cats} top-level)")
    print()

    # Print category tree
    print("Category tree:")
    cursor.execute("""
        SELECT c.Id, c.Name, c.ParentId,
               (SELECT COUNT(*) FROM OfficeDocuments d WHERE d.CategoryId = c.Id) as doc_count
        FROM DocumentCategories c
        ORDER BY c.ParentId IS NOT NULL, c.ParentId, c.SortOrder
    """)
    all_cats = cursor.fetchall()
    children_map = {}
    for cat_id, name, parent_id, doc_count_val in all_cats:
        children_map.setdefault(parent_id, []).append((cat_id, name, doc_count_val))

    def print_tree(parent_id, depth=0):
        for cat_id, name, doc_count_val in children_map.get(parent_id, []):
            indent = "  " * (depth + 1)
            suffix = f" ({doc_count_val} docs)" if doc_count_val > 0 else ""
            print(f"{indent}{name}{suffix}")
            print_tree(cat_id, depth + 1)

    print_tree(None)

    conn.close()


if __name__ == "__main__":
    main()
