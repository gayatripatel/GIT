namespace WebApplication2.DAL
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class TransactionsContext : DbContext
    {
        public TransactionsContext()
            : base("name=TransactionsConnStr")
        {
        }

        public virtual DbSet<testt> testts { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<testt>()
                .Property(e => e.AccountNumber)
                .IsFixedLength();

            modelBuilder.Entity<testt>()
                .Property(e => e.Description)
                .IsUnicode(false);

            modelBuilder.Entity<testt>()
                .Property(e => e.CCode)
                .IsFixedLength();

            modelBuilder.Entity<testt>()
                .Property(e => e.Amount)
                .HasPrecision(19, 4);
        }
    }
}
