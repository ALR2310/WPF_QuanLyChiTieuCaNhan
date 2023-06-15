using LiveCharts;
using LiveCharts.Wpf;
using QuanLyChiTieu.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace QuanLyChiTieu
{
    /// <summary>
    /// Interaction logic for Statics.xaml
    /// </summary>
    public partial class Statics : Window
    {
        public Statics()
        {
            InitializeComponent();
            ShowData();
            TotalPrice();
            TotalDate();
            TotalCount();
        }
        

        void ShowData()     //Hiển thị dữ liệu
        {
            List<STATICS> list = new List<STATICS>();

            string query = "SELECT name, SUM(price) AS price FROM dbo.QLCT GROUP BY name ORDER BY price DESC";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);

            foreach (DataRow item in data.Rows)
            {
                STATICS qlct = new STATICS(item);
                list.Add(qlct);
            }
            dg_list.ItemsSource = list;
        }

        void TotalPrice()       // Hiển thị tổng tiền
        {
            string query = "SELECT SUM(price) AS total FROM dbo.QLCT";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);

            if (data.Rows.Count > 0)
            {
                double totalPrice = Convert.ToDouble(data.Rows[0]["total"]);

                txt_totalprice.Text = string.Format("{0:N0} VNĐ", totalPrice);
            }
        }

        void TotalDate()        //Hiển thị tổng ngày
        {
            string query = "SELECT COUNT(DISTINCT Date) AS total FROM QLCT";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);

            if (data.Rows.Count > 0)
            {
                double totalDate = Convert.ToDouble(data.Rows[0]["total"]);
                string str = "Ngày";
                txt_totaldate.Text = string.Format("{0} {1}", totalDate.ToString(), str);
            }
        }

        void TotalCount()       //Hiển thị tổng số lượng
        {
            string query = "SELECT COUNT(DISTINCT name) AS total FROM dbo.QLCT";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);

            if (data.Rows.Count > 0)
            {
                double totalNameCount = Convert.ToDouble(data.Rows[0]["total"]);
                string str = "Khoản chi";
                txt_totalcount.Text = string.Format("{0} {1}", totalNameCount.ToString(), str);
            }
        }

    }
}
