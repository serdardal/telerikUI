using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;
using WebApplication1.Models.DbModels;

namespace WebApplication1.Data
{
    public class DataContext : DbContext
    {
        public DataContext(DbContextOptions<DataContext> options) : base(options)
        {
            Database.EnsureCreated();
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<CellRecord>()
                .HasIndex(p => new { p.RowIndex, p.ColumnIndex, p.FileName, p.TableIndex }).IsUnique();
        }

        public DbSet<CellRecord> CellRecords { get; set; }

        public DbSet<Default> Defaults { get; set; }
    }
}
