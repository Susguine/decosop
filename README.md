# DecoSOP

A web app for managing Standard Operating Procedures. Organize, search, and edit SOPs from any machine on your network. Runs on a single Windows PC or server — no cloud account needed.

Built with Blazor Server (.NET 10) and SQLite.

## Features

- Nested category hierarchy for organizing SOPs
- Rich text editor for creating and editing documents
- Full-text search
- Favorites — star categories or documents for quick access
- Resizable sidebar with collapsible category tree
- Runs as a Windows Service — starts on boot, always available

## Installation

### Option 1: Download a release (recommended)

1. Go to the [Releases](https://github.com/Susguine/decosop/releases) page and download the latest `.zip` file
2. Extract the zip to a folder on the machine that will host it
3. Open PowerShell **as Administrator**, navigate to the extracted folder, and run:

```powershell
.\install.ps1
```

That's it. The installer will:
- Copy files to `C:\DecoSOP`
- Register DecoSOP as a Windows Service (starts automatically on boot)
- Open port 5098 in Windows Firewall so other machines on the network can connect
- Start the service

When it finishes, it prints the URL to access the app.

**Custom install location or port:**

```powershell
.\install.ps1 -InstallDir "D:\Apps\DecoSOP" -Port 8080
```

### Option 2: Build from source

If you'd rather not run a prebuilt binary, you can build it yourself.

**Prerequisites:**
- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Git

**Steps:**

```powershell
git clone https://github.com/Susguine/decosop.git
cd decosop
dotnet publish -c Release -r win-x64 --self-contained -o ./publish
```

Then copy the install scripts into the output and run the installer:

```powershell
Copy-Item install.ps1, uninstall.ps1 -Destination ./publish/
cd publish
.\install.ps1
```

Or if you prefer to set things up manually:

```powershell
# Copy the publish folder to your install location
Copy-Item -Path ./publish/* -Destination C:\DecoSOP -Recurse -Force

# Create the Windows Service
sc.exe create DecoSOP binPath="C:\DecoSOP\DecoSOP.exe" start=auto

# Open the firewall port
netsh advfirewall firewall add rule name="DecoSOP" dir=in action=allow protocol=TCP localport=5098

# Start the service
Start-Service DecoSOP
```

## Accessing the app

- **On the host machine:** `http://localhost:5098`
- **From other machines on the network:** `http://<host-ip>:5098`

To find the host machine's IP address, run `ipconfig` on it and look for the IPv4 address (usually starts with `192.168.`).

## Updating

### From a release

1. Download the new release zip
2. Extract it and run `install.ps1` again — it will stop the service, update the files (preserving your database), and restart

### From source

```powershell
git pull
dotnet publish -c Release -r win-x64 --self-contained -o ./publish
Stop-Service DecoSOP
Copy-Item -Path ./publish/* -Destination C:\DecoSOP -Recurse -Force -Exclude "decosop.db"
Start-Service DecoSOP
```

## Uninstalling

Run PowerShell **as Administrator**:

```powershell
C:\DecoSOP\uninstall.ps1
```

This removes the service, the firewall rule, and all application files. Your SOP database is deleted too.

To keep your data when uninstalling:

```powershell
C:\DecoSOP\uninstall.ps1 -KeepData
```

## Importing existing documents

If you have SOPs in `.docx`, `.doc`, or `.xlsx` files, the `import_sops.py` script can bulk-import them into the database. It converts documents to HTML and mirrors your folder structure as categories.

**Requires:** Python 3.10+, `mammoth`, `openpyxl`

```bash
pip install mammoth openpyxl
```

Edit `PROTOCOLS_DIR` and `DB_PATH` at the top of the script to point to your files and database, then run:

```bash
python import_sops.py
```

## Development

```bash
dotnet run
```

The app runs at `http://localhost:5098` with hot reload. The SQLite database is created automatically on first run.

### Project structure

```
Components/
  Layout/          - MainLayout, NavMenu (sidebar)
  Pages/           - Home, CategoryView, SopEditor, SopViewer
Data/              - EF Core DbContext
Models/            - Category, SopDocument
Services/          - SopService (data access)
wwwroot/js/        - Editor interop, sidebar resize
```

## Troubleshooting

**"Can't reach the page" from another computer**
- Make sure the firewall rule was created: `netsh advfirewall firewall show rule name="DecoSOP"`
- Verify the service is running: `Get-Service DecoSOP`
- Check that you're using the right IP — run `ipconfig` on the host

**Service starts but immediately stops**
- Run the exe directly to see the error: `C:\DecoSOP\DecoSOP.exe`
- Common cause: another process is already using port 5098. Change the port in `appsettings.json` or reinstall with `-Port 8080`

**Port conflict with another application**
- You can use any port. Reinstall with: `.\install.ps1 -Port 8080`
- Or edit `C:\DecoSOP\appsettings.json` and restart the service

**Need to reset the database**
- Stop the service: `Stop-Service DecoSOP`
- Delete `C:\DecoSOP\decosop.db`
- Start the service: `Start-Service DecoSOP` — a fresh empty database is created on startup
