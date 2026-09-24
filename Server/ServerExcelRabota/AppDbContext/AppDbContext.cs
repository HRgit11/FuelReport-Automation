namespace ServerExcelRabota.AppDbContext
{
    using Microsoft.EntityFrameworkCore;
    using ServerExcelRabota.Model; 

    public class AppDbContextCs : DbContext
    {
        public AppDbContextCs(DbContextOptions<AppDbContextCs> options) : base(options)
        {
        }

        public DbSet<ServerExcelRabota.Model.License> Licenses { get; set; }

    }
}