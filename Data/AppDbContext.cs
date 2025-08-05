using ERPToolsAPI.Models;
using Microsoft.EntityFrameworkCore;
using Tools.Models;

namespace ERPToolsAPI.Data
{
    public class ERPToolsDbContext : DbContext
    {
        public ERPToolsDbContext(DbContextOptions<ERPToolsDbContext> options) : base(options) { }

       // public DbSet<ToolRecord> ToolRecords { get; set; }
        public DbSet<UserLoginLog> UserLoginLogs { get; set; }
        public DbSet<ExcelUpload> ExcelUploads { get; set; }
        public DbSet <BoxCapacity> BoxCapacity { get; set; }
        public DbSet <EnvelopeType> EnvelopesTypes { get; set; }
        public DbSet <NRData> NRDatas { get; set; }
        public DbSet <Field> Fields { get; set; }
        public DbSet <ProjectConfig> ProjectConfigs { get; set; }
        public DbSet <ToolRecord> ToolRecords { get; set; }
        public DbSet <Module> Modules { get; set; }
        public DbSet <ToolConfig> ToolConfigs { get; set; }
     
    }

    
}
