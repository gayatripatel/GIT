namespace WebApplication2.DAL
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("testt")]
    public partial class testt
    {
        public int Id { get; set; }

        [Required]
        [StringLength(10)]
        public string AccountNumber { get; set; }

        [Required]
        public string Description { get; set; }

        [Required]
        [StringLength(10)]
        public string CCode { get; set; }

        [Column(TypeName = "money")]
        public decimal Amount { get; set; }
    }
}
