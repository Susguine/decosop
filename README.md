# DecoSOP

Internal SOP management app for the dental office. Staff can browse, search, create, and edit standard operating procedures through a web interface accessible from any machine on the network.

Built with Blazor Server (.NET 10) and SQLite.

## Features

- Nested category hierarchy for organizing SOPs
- Rich text editor (Jodit) for creating and editing documents
- Full-text search from the sidebar
- Favorites — star categories or documents for quick access
- Resizable sidebar with collapsible category tree
- Runs as a Windows Service on the office server

## Running locally

```
dotnet run
```

App will be available at `http://localhost:5098`. The SQLite database (`decosop.db`) is created automatically on first run.

## Deploying to the server

Publish a self-contained build:

```
dotnet publish -c Release -r win-x64 --self-contained -o ./publish
```

Copy the `publish/` folder to the server (e.g. `C:\DecoSOP`), then register it as a Windows Service:

```powershell
sc.exe create DecoSOP binPath="C:\DecoSOP\DecoSOP.exe" start=auto
sc.exe start DecoSOP
```

Make sure port 5098 is open in Windows Firewall:

```powershell
netsh advfirewall firewall add rule name="DecoSOP" dir=in action=allow protocol=TCP localport=5098
```

Employees can then access it at `http://<server-ip>:5098`.

## Importing SOPs

The `import_sops.py` script bulk-imports `.docx`, `.doc`, and `.xlsx` files from a directory into the database. It mirrors the folder structure as nested categories.

Requires Python 3.10+ with `mammoth` and `openpyxl`:

```
pip install mammoth openpyxl
python import_sops.py
```

Edit the `PROTOCOLS_DIR` and `DB_PATH` variables at the top of the script to point to your files and database.

## Project structure

```
Components/
  Layout/          - MainLayout, NavMenu (sidebar)
  Pages/           - Home, CategoryView, SopEditor, SopViewer
Data/              - EF Core DbContext
Models/            - Category, SopDocument
Services/          - SopService (all data access)
wwwroot/js/        - Jodit editor interop, sidebar resize
```
