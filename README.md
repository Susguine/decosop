# DecoSOP

Internal web application for managing and viewing Standard Operating Procedures (SOPs) for dental office staff.

## Tech Stack

- **Blazor Web App** (.NET 10, Interactive Server)
- **SQLite** via Entity Framework Core (local database, no external server needed)
- **Markdig** for Markdown rendering

## Features

- Browse SOPs organized by category in a sidebar directory
- View SOPs rendered from Markdown with clean formatting
- Create, edit, and delete SOPs with a built-in Markdown editor with live preview
- Manage categories (add, rename, reorder, delete)
- Full-text search across all SOPs

## Getting Started

### Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)

### Run Locally

```bash
dotnet run
```

The app will be available at `http://localhost:5000`.

### Database

The SQLite database (`decosop.db`) is created automatically on first run via EF Core migrations. No external database server needed.

To apply migrations manually:

```bash
dotnet ef database update
```

## Project Structure

```
DecoSOP/
├── Components/
│   ├── Layout/          # Main layout, sidebar navigation
│   └── Pages/           # Blazor pages (Home, Viewer, Editor)
├── Data/                # EF Core DbContext
├── Models/              # Category, SopDocument entities
├── Services/            # Business logic (SopService)
├── wwwroot/css/         # Stylesheets
└── Program.cs           # App startup and DI configuration
```
