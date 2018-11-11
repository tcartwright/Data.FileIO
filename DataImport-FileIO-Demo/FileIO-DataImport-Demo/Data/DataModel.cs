using System.Data.Entity;

namespace FileIODemo.Data
{
    public partial class DataModel : DbContext
	{
		public DataModel()
			: base("name=DataModel")
		{
		}

		public virtual DbSet<Company> Company { get; set; }

		protected override void OnModelCreating(DbModelBuilder modelBuilder)
		{
			modelBuilder.Entity<Company>()
				.Property(e => e.LegalName)
				.IsUnicode(false);

			modelBuilder.Entity<Company>()
				.Property(e => e.DBAName)
				.IsUnicode(false);

			modelBuilder.Entity<Company>()
				.Property(e => e.UserId)
				.IsUnicode(false);

		}
	}
}
