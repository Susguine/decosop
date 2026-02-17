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
            e.HasIndex(c => c.Name).IsUnique();
            e.HasMany(c => c.Documents)
             .WithOne(d => d.Category)
             .HasForeignKey(d => d.CategoryId)
             .OnDelete(DeleteBehavior.Cascade);
        });

        modelBuilder.Entity<SopDocument>(e =>
        {
            e.HasIndex(d => new { d.CategoryId, d.Title }).IsUnique();
        });

        // Seed sample data
        modelBuilder.Entity<Category>().HasData(
            new Category { Id = 1, Name = "Front Desk", SortOrder = 0 },
            new Category { Id = 2, Name = "Clinical", SortOrder = 1 },
            new Category { Id = 3, Name = "Billing & Insurance", SortOrder = 2 }
        );

        modelBuilder.Entity<SopDocument>().HasData(
            new SopDocument
            {
                Id = 1,
                CategoryId = 1,
                Title = "Patient Check-In",
                SortOrder = 0,
                CreatedAt = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                UpdatedAt = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                MarkdownContent = """
                # Patient Check-In Procedure

                ## Purpose
                Ensure a smooth and welcoming check-in experience for every patient.

                ## Steps

                1. **Greet the patient** by name when they approach the front desk
                2. **Verify identity** — confirm date of birth and address
                3. **Check insurance** — verify current insurance card is on file
                4. **Update forms** — ask if any personal or medical information has changed
                5. **Collect copay** if applicable
                6. **Notify clinical staff** that the patient has arrived

                ## Notes
                - If the patient is more than 15 minutes late, check with the provider before seating
                - New patients should be given the welcome packet and asked to complete intake forms
                """
            },
            new SopDocument
            {
                Id = 2,
                CategoryId = 1,
                Title = "Answering the Phone",
                SortOrder = 1,
                CreatedAt = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                UpdatedAt = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                MarkdownContent = """
                # Answering the Phone

                ## Greeting
                > "Thank you for calling [Practice Name], this is [Your Name]. How can I help you today?"

                ## Key Guidelines
                - Answer within **3 rings**
                - Always be **friendly and professional**
                - If placing a caller on hold, ask permission first
                - Take a detailed message if the intended person is unavailable

                ## Common Call Types
                | Call Type | Action |
                |-----------|--------|
                | Schedule appointment | Open scheduling module |
                | Cancel/reschedule | Check cancellation policy |
                | Billing question | Transfer to billing |
                | Emergency | Assess urgency, notify provider immediately |
                """
            },
            new SopDocument
            {
                Id = 3,
                CategoryId = 2,
                Title = "Sterilization Protocol",
                SortOrder = 0,
                CreatedAt = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                UpdatedAt = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                MarkdownContent = """
                # Instrument Sterilization Protocol

                ## Purpose
                Maintain infection control standards per OSHA and CDC guidelines.

                ## Steps

                1. **Transport** — carry contaminated instruments to sterilization area in a covered container
                2. **Pre-clean** — rinse instruments under running water to remove visible debris
                3. **Ultrasonic cleaning** — place instruments in ultrasonic cleaner for the recommended cycle time
                4. **Rinse and dry** — rinse thoroughly, then dry with lint-free cloth
                5. **Package** — place instruments in sterilization pouches, seal properly
                6. **Autoclave** — run appropriate sterilization cycle (refer to autoclave manual for settings)
                7. **Verify** — check chemical indicator strip after cycle completes
                8. **Store** — place sealed pouches in designated clean storage area

                ## Spore Testing
                - Perform biological spore testing **weekly**
                - Log results in the sterilization log binder
                """
            }
        );
    }
}
