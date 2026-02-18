using DecoSOP.Components;
using DecoSOP.Data;
using DecoSOP.Services;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

builder.Host.UseWindowsService();
builder.WebHost.UseUrls("http://0.0.0.0:5098");

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Use absolute path so the DB is always next to the app (not in System32 when running as a service)
var dbPath = Path.Combine(AppContext.BaseDirectory, "decosop.db");
builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseSqlite($"Data Source={dbPath}"));

builder.Services.AddScoped<SopService>();

var app = builder.Build();

// Auto-create/migrate the database on startup
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();

    // Add IsFavorited columns to existing databases (EnsureCreated only creates new DBs)
    var conn = db.Database.GetDbConnection();
    await conn.OpenAsync();
    try
    {
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
    }
    finally
    {
        await conn.CloseAsync();
    }
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

app.Run();
