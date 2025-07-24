using ERPToolsAPI.Models;
using Microsoft.EntityFrameworkCore;

namespace ERPToolsAPI.Data
{
    public class ERPToolsDbContext : DbContext
    {
        public ERPToolsDbContext(DbContextOptions<ERPToolsDbContext> options) : base(options) { }

       // public DbSet<ToolRecord> ToolRecords { get; set; }
        public DbSet<UserLoginLog> UserLoginLogs { get; set; }
        public DbSet<ExcelUpload> ExcelUploads { get; set; }
    }

    
}
