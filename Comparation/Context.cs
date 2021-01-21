using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace Comparation
{
    public class Context : DbContext
    {
        protected override void OnConfiguring(DbContextOptionsBuilder options)
            => options.UseSqlite("Data Source=comparationdb.db");

        public DbSet<Table1> Table1 { get; set; }
        public DbSet<Table2> Table2 { get; set; }
    }
}
