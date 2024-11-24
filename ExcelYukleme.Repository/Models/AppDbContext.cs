using Microsoft.EntityFrameworkCore;

namespace ExcelYukleme.Repository.Models
{
    public class AppDbContext :DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
       
        public DbSet<PersonelModel> Personeller { get; set; }
        public DbSet<KodIlcelerModel> KodIlceler { get; set; }
        public DbSet<RutbelerModel> KodRutbeler { get; set; }
        public DbSet<BirimlerModel> KodBirimler { get; set; }
        public DbSet<KanGruplariModel> KodKanGruplari { get; set; }
        public DbSet<PersonelRolleriModel>PersonelRolleri { get; set; }
    }

}
