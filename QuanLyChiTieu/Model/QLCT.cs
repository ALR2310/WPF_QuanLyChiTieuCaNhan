using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyChiTieu.Model
{
    public class QLCT
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public double Price { get; set; }
        public string MoreInfo { get; set; }
        public TimeSpan Time { get; set; }

        public QLCT(int id, DateTime date, string name, float price, string moreInfo, TimeSpan time)
        {
            this.Id = id;
            this.Date = date;
            this.Name = name;
            this.Price = price;
            this.MoreInfo = moreInfo;
            this.Time = time;
        }

        public QLCT(DataRow row)
        {
            this.Id = (int)row["id"];
            this.Date = (DateTime)row["date"];
            this.Name = row["name"].ToString();
            this.Price = (double)row["price"];
            this.MoreInfo = row["moreInfo"].ToString();
            this.Time = (TimeSpan)row["time"];
        }
    }

}
