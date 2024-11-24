using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelYukleme.Repository.Models
{
    public class BirimlerModel
    {
        public int Id { get; set; }
        public string Ad { get; set; }
        public int MulkiBirimId { get; set; }
        public bool Iptalmi { get; set; }
    }
}
