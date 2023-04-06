using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab_data.Model
{
    public  class ApplicationContext : DbContext
    {
        public DbSet<employee> employees { get; set; } = null!;
        public DbSet<Post> Posts { get; set; } = null!;
        public DbSet<children> childrens { get; set; } = null!;
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite("Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "/Data/laboratory.db");
        }
    }
}
