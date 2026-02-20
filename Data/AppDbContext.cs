using DecoSOP.Models;
using Microsoft.EntityFrameworkCore;

namespace DecoSOP.Data;

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    public DbSet<Category> Categories => Set<Category>();
    public DbSet<SopDocument> Documents => Set<SopDocument>();
    public DbSet<DocumentCategory> DocumentCategories => Set<DocumentCategory>();
    public DbSet<OfficeDocument> OfficeDocuments => Set<OfficeDocument>();
    public DbSet<WebDocCategory> WebDocCategories => Set<WebDocCategory>();
    public DbSet<WebDocument> WebDocuments => Set<WebDocument>();
    public DbSet<SopCategory> SopCategories => Set<SopCategory>();
    public DbSet<SopFile> SopFiles => Set<SopFile>();

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

        modelBuilder.Entity<DocumentCategory>(e =>
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

        modelBuilder.Entity<OfficeDocument>(e =>
        {
            e.HasIndex(d => new { d.CategoryId, d.Title }).IsUnique();
        });

        modelBuilder.Entity<WebDocCategory>(e =>
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

        modelBuilder.Entity<WebDocument>(e =>
        {
            e.HasIndex(d => new { d.CategoryId, d.Title }).IsUnique();
        });

        modelBuilder.Entity<SopCategory>(e =>
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

        modelBuilder.Entity<SopFile>(e =>
        {
            e.HasIndex(d => new { d.CategoryId, d.Title }).IsUnique();
        });
    }
}
