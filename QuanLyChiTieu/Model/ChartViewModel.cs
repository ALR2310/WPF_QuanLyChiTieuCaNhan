using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyChiTieu.Model
{
    public class ChartViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<STATICS> _data;

        public ObservableCollection<STATICS> Data
        {
            get { return _data; }
            set
            {
                _data = value;
                OnPropertyChanged(nameof(Data));
            }
        }

        public ChartViewModel()
        {
            LoadData();
        }

        private void LoadData()
        {
            // Thực hiện truy vấn và lấy dữ liệu từ database
            string query = "SELECT name, SUM(price) AS price FROM dbo.QLCT GROUP BY name";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);

            // Khởi tạo danh sách dữ liệu
            ObservableCollection<STATICS> dataList = new ObservableCollection<STATICS>();

            // Thêm các bản ghi vào danh sách dữ liệu
            foreach (DataRow item in data.Rows)
            {
                STATICS charts = new STATICS(item);
                dataList.Add(charts);
            }

            // Gán danh sách dữ liệu cho thuộc tính Data
            Data = dataList;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

}
