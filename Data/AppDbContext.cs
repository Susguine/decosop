using DecoSOP.Models;
using Microsoft.EntityFrameworkCore;

namespace DecoSOP.Data;

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    public DbSet<Category> Categories => Set<Category>();
    public DbSet<SopDocument> Documents => Set<SopDocument>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Category>(e =>
        {
            e.HasIndex(c => new { c.ParentId, c.Name }).IsUnique();
            e.HasMany(c => c.Children)
             .WithOne(c => c.Parent)
             .HasForeignKey(c => c.ParentId)
             .OnDelete(DeleteBehavior.Restrict);
            e.HasMany(c => c.Documents)
             .WithOne(d => d.Category)
             .HasForeignKey(d => d.CategoryId)
             .OnDelete(DeleteBehavior.Cascade);
        });

        modelBuilder.Entity<SopDocument>(e =>
        {
            e.HasIndex(d => new { d.CategoryId, d.Title }).IsUnique();
        });

        // Data is populated via import_sops.py script
    }
}
