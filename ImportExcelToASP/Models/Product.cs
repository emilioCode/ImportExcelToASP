using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ImportExcelToASP.Models
{
    public class Product
    {
        public int ProductId { get; set; }
        public string Nombre { get; set; }
        public decimal Precio { get; set; }
        public int Cantidad { get; set; }
    }
}