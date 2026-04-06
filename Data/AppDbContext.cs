using ERPToolsAPI.Models;
using Microsoft.EntityFrameworkCore;
using Tools.Models;

namespace ERPToolsAPI.Data
{
    public class ERPToolsDbContext : DbContext
    {
        public ERPToolsDbContext(DbContextOptions<ERPToolsDbContext> options) : base(options) { }
        public DbSet<UserLoginLog> UserLoginLogs { get; set; }
        public DbSet<ExcelUpload> ExcelUploads { get; set; }
        public DbSet <BoxCapacity> BoxCapacity { get; set; }
        public DbSet <EnvelopeType> EnvelopesTypes { get; set; }
        public DbSet <NRData> NRDatas { get; set; }

        public DbSet<ChangedNRData> ChangedNRData { get; set; }
        public DbSet <Field> Fields { get; set; }
        public DbSet <ProjectConfig> ProjectConfigs { get; set; }
        public DbSet<MProjectConfigs> MProjectConfigs { get; set; }

        public DbSet <ToolRecord> ToolRecords { get; set; }
        public DbSet <Module> Modules { get; set; }
        public DbSet <ToolConfig> ToolConfigs { get; set; }
        public DbSet<ExtrasConfiguration> ExtraConfigurations { get; set; }
        public DbSet<MExtraConfigurations> MExtraConfigurations { get; set; }
        public DbSet<ExtraEnvelopes> ExtrasEnvelope { get; set; }
        public DbSet<ExtraType> ExtraType { get; set; }
        public DbSet<EnvelopeBreakage> EnvelopeBreakages { get; set; }
        public DbSet<Project> Projects { get; set; }
        public DbSet<ErrorLog> ErrorLogs { get; set; }
        public DbSet<EventLog> EventLogs { get; set; }
        public DbSet<ConflictingFields> ConflictingFields { get; set; }
        public DbSet<BoxBreakingResult> BoxBreakingResults { get; set; }
        public DbSet<EnvelopeBreakingResult> EnvelopeBreakingResults { get; set; }
        public DbSet<Mss> Mss { get; set; }
        public DbSet<RPTTemplate> RPTTemplates { get; set; }

        public DbSet<RPTMapping> RPTMappings { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<ExtraEnvelopes>()
                .HasIndex(e => new { e.CatchNo, e.ExtraId, e.ProjectId})
                .IsUnique();
            modelBuilder.Entity<EnvelopeBreakage>()
                .HasIndex(e => new { e.NrDataId, e.ProjectId })
                .IsUnique();
            modelBuilder.Entity<ExtrasConfiguration>()
                .HasIndex(e => new { e.ExtraType, e.ProjectId })
                .IsUnique();
            modelBuilder.Entity<ProjectConfig>()
                .HasIndex(e => e.ProjectId)
                .IsUnique();
            modelBuilder.Entity<Project>()
                .HasIndex(e => e.ProjectId)
                .IsUnique();
            modelBuilder.Entity<BoxBreakingResult>()
                .HasIndex(e => new { e.ProjectId, e.BoxNo,e.UploadBatch })
                .IsUnique();
        }
        
    }

    
}
