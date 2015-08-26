using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel.DataAnnotations;


namespace WebApplication2.Models
{
    public class Transaction
    {
        [Key]
        public int Id { get; private set; }

        public string Account { get; set; }

        public string Description { get; set; }

        public string CurrencyCode { get; set; }

        public decimal Amount { get; set; }
    }
}
