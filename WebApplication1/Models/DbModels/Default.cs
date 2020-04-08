using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models.DbModels
{
    public class Default
    {
        [Key]
        public Guid Id { get; set; }

        public string Key { get; set; }

        public string Value { get; set; }
    }
}
