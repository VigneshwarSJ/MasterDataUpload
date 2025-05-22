using Microsoft.EntityFrameworkCore;

namespace BackEnd.Models
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<ExcelFileInfo> ExcelFiles { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Configure ExcelFileInfo entity
            modelBuilder.Entity<ExcelFileInfo>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.FileName).IsRequired();
                entity.Property(e => e.UploadDate).IsRequired();
                entity.Property(e => e.Status).IsRequired();
            });
        }
    }
} 