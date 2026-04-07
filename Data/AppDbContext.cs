using ERPToolsAPI.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System.Linq;
using System.Text.Json;
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
<<<<<<< HEAD
                .HasIndex(e => new { e.ProjectId, e.BoxNo })
                .IsUnique(false);

            var uploadListConverter = new ValueConverter<List<int>, string>(
                v => SerializeUploadList(v),
                v => DeserializeUploadList(v)
            );

            var uploadListComparer = new ValueComparer<List<int>>(
                (l, r) => (l ?? new List<int>()).SequenceEqual(r ?? new List<int>()),
                v => (v ?? new List<int>()).Aggregate(0, (a, item) => HashCode.Combine(a, item.GetHashCode())),
                v => (v ?? new List<int>()).ToList()
            );

            modelBuilder.Entity<NRData>()
                .Property(e => e.UploadList)
                .HasConversion(uploadListConverter)
                .Metadata.SetValueComparer(uploadListComparer);
        }

        private static List<int> DeserializeUploadList(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return new List<int>();
            }

            try
            {
                using var doc = JsonDocument.Parse(value);
                var root = doc.RootElement;
                if (root.ValueKind == JsonValueKind.Array)
                {
                    return root
                        .EnumerateArray()
                        .Where(x => x.ValueKind == JsonValueKind.Number)
                        .Select(x => x.GetInt32())
                        .ToList();
                }
                if (root.ValueKind == JsonValueKind.Number)
                {
                    return new List<int> { root.GetInt32() };
                }
                if (root.ValueKind == JsonValueKind.String)
                {
                    var text = root.GetString() ?? string.Empty;
                    return ParseUploadListText(text);
                }
            }
            catch
            {
                // Fallback to plain text parsing
            }

            return ParseUploadListText(value);
        }

        private static List<int> ParseUploadListText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return new List<int>();
            }
            var parts = value
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(v => v.Trim())
                .Where(v => int.TryParse(v, out _))
                .Select(int.Parse)
                .ToList();

            return parts;
        }

        private static string SerializeUploadList(List<int> value)
        {
            return JsonSerializer.Serialize(value ?? new List<int>());
=======
                .HasIndex(e => new { e.ProjectId, e.BoxNo,e.UploadBatch })
                .IsUnique();
>>>>>>> Production
        }
        
    }

    
}
