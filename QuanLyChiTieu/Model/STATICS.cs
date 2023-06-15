using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyChiTieu.Model
{
    public class STATICS
    {
        public string Name { get; set; }
        public double Price { get; set; }

        public STATICS(string name, float price)
        {
            this.Name = name;
            this.Price = price;
        }

        public STATICS(DataRow row)
        {
            this.Name = row["name"].ToString();
            this.Price = (double)row["price"];
        }
    }
}
