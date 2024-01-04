using Microsoft.EntityFrameworkCore;

namespace ExcelApi.Models
{
    public class Aplicationdbcontext:DbContext
    {
        public Aplicationdbcontext(DbContextOptions<Aplicationdbcontext> optipns) : base(optipns)
        {

        }

        public  DbSet<testtable> tests {  get; set; }
    }
}
