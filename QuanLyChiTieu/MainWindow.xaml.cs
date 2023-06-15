using iTextSharp.text.pdf;
using iTextSharp.text;
using QuanLyChiTieu.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml.Linq;
using System.Collections;
using Microsoft.Win32;
using OfficeOpenXml.Table;
using OfficeOpenXml;

namespace QuanLyChiTieu
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ShowData();     //gọi hàm load khi chương trình khởi động
            UpdTimer();     //bắt đầu đếm giờ khi chương trình khởi động
            LoadCombobox();     //load dữ liệu của cbbox khi chương trình khởi động
            ScrollDataGrid();   //cuộn data xuống cuối khi chương trình khởi động
            TotalPrice();       //Hiển thị tổng tiền khi chương trình khởi động
        }

        #region Hàm

        void ShowData()     //Hiển thị dữ liệu
        {
            List<QLCT> list = new List<QLCT>();

            string query = "SELECT * FROM dbo.QLCT ORDER BY date ASC";
            try
            {
                DataTable data = DataProvider.Instance.ExecuteQuery(query);

                foreach (DataRow item in data.Rows)
                {
                    QLCT qlct = new QLCT(item);
                    list.Add(qlct);
                }
                dg_show.ItemsSource = list;
            }
            catch
            {
                MessageBox.Show("Có Lỗi Khi Hiển Thị Dữ Liệu\nVui Lòng Kiểm Tra Lại", "Thông Báo", MessageBoxButton.OK);
            }
        }

        void InsertData(DateTime date, string name, double price, string moreInfo, TimeSpan time)       //Thêm dữ liệu
        {
            string query = "EXEC dbo.sp_InsertData @date , @name , @price , @moreInfo , @time ";
            try
            {
                DataProvider.Instance.ExecuteNonQuery(query, new object[] { date, name, price, moreInfo, time });
                ShowData();
            }
            catch
            {
                MessageBox.Show("Có Lỗi Khi Thêm Dữ Liệu\nVui Lòng Kiểm Tra Lại", "Thông Báo", MessageBoxButton.OK);
            }
        }

        void UpdateData(int id, DateTime date, string name, double price, string moreInfo, TimeSpan time)       //Cập nhật dữ liệu
        {
            string query = "EXEC dbo.sp_UpdateData @id , @date , @name , @price , @moreInfo , @time ";
            try
            {
                DataProvider.Instance.ExecuteNonQuery(query, new object[] { id, date, name, price, moreInfo, time });
                ShowData();
            }
            catch
            {
                MessageBox.Show("Có Lỗi Khi Cập Nhật Dữ Liệu\nVui Lòng Kiểm Tra Lại", "Thông Báo", MessageBoxButton.OK);
            }
        }

        void DeleteData(int id)         //Xoá dữ liệu
        {
            string query = "EXEC dbo.sp_DeleteData @id ";
            try
            {
                DataProvider.Instance.ExecuteNonQuery(query, new object[] { id });
                ShowData();
            }
            catch
            {
                MessageBox.Show("Có Lỗi Khi Xoá Dữ Liệu\nVui Lòng Kiểm Tra Lại", "Thông Báo", MessageBoxButton.OK);
            }

        }

        void ClearData()    //Dọn dữ liệu các textbox
        {
            txt_ID.Clear();
            pk_Date.SelectedDate = DateTime.Now;
            cb_name.SelectedIndex = -1;
            txt_Price.Clear();
            txt_moreInfo.Clear();
            tp_time.SelectedTime = DateTime.Now;
        }

        void UpdTimer()     //Cập nhật thời gian của TimerPicker liên tục
        {
            timer = new DispatcherTimer();
            timer.Tick += new EventHandler(timer_Tick);
            timer.Interval = new TimeSpan(0, 0, 1);
            timer.Start();
        }

        private DispatcherTimer timer;

        private void timer_Tick(object sender, EventArgs e)
        {
            tp_time.SelectedTime = DateTime.Today + DateTime.Now.TimeOfDay;
        }

        void TextCap(UIElement uiElement)       //Viết hoa chữ cái đầu của Textbox và Combobox
        {
            // Lấy thông tin ngôn ngữ mặc định của hệ thống
            CultureInfo cultureInfo = CultureInfo.CurrentCulture;

            // Lấy đối tượng TextInfo của ngôn ngữ đó
            TextInfo textInfo = cultureInfo.TextInfo;

            // Kiểm tra nếu uiElement là TextBox và TextBox đó không rỗng
            if (uiElement is TextBox textBox && !string.IsNullOrEmpty(textBox.Text))
            {
                // Chuyển đổi chuỗi thành dạng tiêu đề
                string titleCaseString = textInfo.ToTitleCase(textBox.Text);

                // Gán lại chuỗi đã được chuyển đổi vào TextBox
                textBox.Text = titleCaseString;
            }
            // Kiểm tra nếu uiElement là ComboBox và ComboBox đó không rỗng
            else if (uiElement is ComboBox comboBox && comboBox.SelectedItem != null)
            {
                // Chuyển đổi chuỗi thành dạng tiêu đề
                string titleCaseString = textInfo.ToTitleCase(comboBox.SelectedItem.ToString());

                // Gán lại chuỗi đã được chuyển đổi vào ComboBox
                comboBox.SelectedItem = titleCaseString;
            }
        }

        void LoadCombobox()     //Hiển thị dữ liệu lên combobox
        {
            List<QLCT> list = new List<QLCT>();

            string query = "select * from dbo.qlct";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);
            foreach (DataRow item in data.Rows)
            {
                QLCT qlct = new QLCT(item);
                list.Add(qlct);
            }
            cb_name.ItemsSource = list;
            cb_name.DisplayMemberPath = "Name";
        }

        void ScrollDataGrid()   //Cuộn xuống cuối của Datagid
        {
            // Chọn hàng cuối cùng trong DataGrid
            int lastRowIndex = dg_show.Items.Count - 1;
            dg_show.SelectedIndex = lastRowIndex;

            // Scroll đến hàng cuối cùng được chọn
            dg_show.ScrollIntoView(dg_show.SelectedItem);
        }

        void CalculateTotal()   //Tính tổng số tiền khi nhấn vào 1 hàng trên datagrid
        {
            double totalPrice = 0;

            // Lấy tên khoản chi được chọn
            QLCT selectedItem = dg_show.SelectedItem as QLCT;
            if (selectedItem != null)
            {
                string selectedName = selectedItem.Name;

                // Tính tổng số tiền của các khoản chi có tên tương ứng
                foreach (QLCT item in dg_show.Items)
                {
                    if (item.Name == selectedName)
                    {
                        totalPrice += item.Price;
                    }
                }

                // Gán giá trị kết quả cho TextBox
                txt_totalprice.Text = totalPrice.ToString("N0") + " VNĐ";
            }
        }

        private int CountItem(string columnName, string itemValue)  //đếm số lượng trong datagrid
        {
            int count = 0;
            foreach (QLCT item in dg_show.Items)
            {
                if (item.Name == itemValue)
                {
                    count++;
                }
            }
            return count;
        }

        void CountDataItem()    //đếm số lượng trong datagrid
        {
            var selectedItem = dg_show.SelectedItem as QLCT;
            if (selectedItem != null)
            {
                // Lấy tên của hàng đang được chọn
                string selectedName = selectedItem.Name;

                // Đếm số lượng item có cùng tên với hàng đang được chọn
                int count = CountItem("Name", selectedName);

                // Hiển thị kết quả lên textbox
                txt_count.Text = count.ToString();
            }
        }

        void SearchDataText()   //Tìm kiếm theo văn bản
        {
            string search = txt_Search.Text.Trim();
            string dateFilter = dp_search.SelectedDate?.ToString("yyyy-MM-dd");

            if (string.IsNullOrEmpty(search) && string.IsNullOrEmpty(dateFilter))
            {
                ShowData();
            }
            else
            {
                try
                {
                    string query = "SELECT * FROM dbo.QLCT WHERE 1 = 1";
                    if (!string.IsNullOrEmpty(search))
                    {
                        query += " AND name LIKE N'%" + search + "%'";
                    }
                    if (!string.IsNullOrEmpty(dateFilter))
                    {
                        query += " AND CONVERT(DATE, Date, 23) = '" + dateFilter + "'";
                    }
                    DataTable data = DataProvider.Instance.ExecuteQuery(query);

                    List<QLCT> list = new List<QLCT>();
                    foreach (DataRow item in data.Rows)
                    {
                        QLCT qlct = new QLCT(item);
                        list.Add(qlct);
                    }
                    dg_show.ItemsSource = list;
                }
                catch
                {
                    MessageBox.Show("Có lỗi khi tìm kiếm\nVui lòng kiểm tra lại", "Thông Báo", MessageBoxButton.OK);
                }
            }
        }

        void SearchDateSeleted()    //Tìm kiếm theo ngày
        {
            DateTime? selectedDate = dp_search.SelectedDate;

            if (selectedDate != null)
            {
                string search = selectedDate.Value.ToString("yyyy-MM-dd");

                string query = "SELECT * FROM dbo.QLCT WHERE date = '" + search + "'";
                DataTable data = DataProvider.Instance.ExecuteQuery(query);

                List<QLCT> list = new List<QLCT>();
                foreach (DataRow item in data.Rows)
                {
                    QLCT qlct = new QLCT(item);
                    list.Add(qlct);
                }
                dg_show.ItemsSource = list;
            }
            else
            {
                ShowData();
            }
        }

        void TotalPrice()       // Hiển thị tổng tiền
        {
            string query = "SELECT SUM(price) AS total FROM dbo.QLCT";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);

            if (data.Rows.Count > 0)
            {
                double totalPrice = Convert.ToDouble(data.Rows[0]["total"]);

                txt_total.Text = string.Format("{0:N0} VNĐ", totalPrice);
            }
        }

        void ExportToPdf(DataGrid dg)    //Xuất File PDF
        {
            // Tạo SaveFileDialog để lấy đường dẫn và tên tệp để lưu
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf|All files (*.*)|*.*";

            // Nếu người dùng chọn OK thì tiếp tục thực hiện tạo tệp PDF
            if (saveFileDialog.ShowDialog() == true)
            {
                // Tạo một tệp PDF mới
                Document document = new Document();
                PdfWriter.GetInstance(document, new FileStream(saveFileDialog.FileName, FileMode.Create));
                document.Open();

                // Thêm font chữ mới vào tài liệu PDF
                BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\Arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                Font font = new Font(baseFont, 12, Font.NORMAL, BaseColor.BLACK);

                // Tạo một bảng PDF và thêm các cột tương ứng với cột trong DataGrid
                PdfPTable pdfTable = new PdfPTable(dg.Columns.Count);
                foreach (DataGridColumn column in dg.Columns)
                {
                    pdfTable.AddCell(new Phrase(column.Header.ToString()));
                }

                // Thêm các dòng dữ liệu từ DataGrid vào bảng PDF
                foreach (object item in dg.Items)
                {
                    if (item != null)
                    {
                        foreach (DataGridColumn column in dg.Columns)
                        {
                            if (column.GetCellContent(item) is TextBlock)
                            {
                                TextBlock cellContent = column.GetCellContent(item) as TextBlock;
                                pdfTable.AddCell(new Phrase(cellContent.Text));
                            }
                        }
                    }
                }

                // Thêm bảng PDF vào tài liệu và đóng tài liệu
                document.Add(pdfTable);
                document.Close();

                // Hiển thị thông báo khi xuất tệp PDF thành công
                MessageBox.Show("Xuất File PDF Thành Công!", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        void ExportToExcel()     //Xuất file Excel
        {
            // Khởi tạo SaveFileDialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                // Lấy đường dẫn file Excel từ SaveFileDialog
                string filePath = saveFileDialog.FileName;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Khởi tạo ExcelPackage
                using (ExcelPackage package = new ExcelPackage())
                {
                    // Tạo một worksheet mới
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    // Lấy dữ liệu từ DataGrid và đổ vào worksheet
                    for (int i = 0; i < dg_show.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dg_show.Columns[i].Header;
                        for (int j = 0; j < dg_show.Items.Count; j++)
                        {
                            var cellValue = dg_show.Columns[i].GetCellContent(dg_show.Items[j])?.ToString();
                            worksheet.Cells[j + 2, i + 1].Value = cellValue;
                        }
                    }

                    // Format worksheet
                    worksheet.Cells.AutoFitColumns();
                    ExcelTable table = worksheet.Tables.Add(new ExcelAddressBase(1, 1, dg_show.Items.Count + 1, dg_show.Columns.Count), "Table1");
                    table.TableStyle = TableStyles.Light9;

                    // Lưu ExcelPackage xuống file
                    package.SaveAs(new FileInfo(filePath));
                }

                MessageBox.Show("Xuất dữ liệu ra file Excel thành công!", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        #endregion

        #region Sự Kiện

        private void bt_Add_Click(object sender, RoutedEventArgs e)     //Thêm dữ liệu
        {
            if (string.IsNullOrEmpty(txt_moreInfo.Text))    //Nếu info rỗng thì ghi giá trị mặt định
            {
                txt_moreInfo.Text = "Không có thông tin";
            }
            TextCap(cb_name);   //Gọi hàm viết hoa
            txt_Price.Text = string.Concat(txt_Price.Text, "000");  //Thêm 000 vào cuối chuỗi
            try     //bắt lỗi nếu người dùng nhập các ký tự đặt biệt vào các textbox
            {
                DateTime date = pk_Date.SelectedDate ?? DateTime.Now;   // Nếu người dùng không chọn ngày thì lấy ngày hiện tại
                string name = cb_name.Text;
                double price = double.Parse(txt_Price.Text);
                string moreInfo = txt_moreInfo.Text;
                TimeSpan time = TimeSpan.Parse(tp_time.Text);

                InsertData(date, name, price, moreInfo, time);
                LoadCombobox();     //sau khi thêm mới sẽ tải lại cbbox
                TotalPrice();       //sau khi thêm mới sẽ tải lại tổng tiền
            }
            catch
            {
                MessageBox.Show("Có Lỗi Khi Thêm Dữ Liệu\nVui Lòng Kiểm Tra Lại", "Thông Báo", MessageBoxButton.OK);
            }
            ScrollDataGrid();   //cuộn xuống mục cuối của datagird

            ClearData();    //Dọn dữ liệu sau khi thêm mới
            cb_name.Focus();    //nhắm chuột vào ô textbox này
        }

        private void bt_Upd_Click(object sender, RoutedEventArgs e)     //Cập nhật dữ liệu
        {
            TextCap(cb_name);   //Gọi hàm viết hoa
            try     //bắt lỗi nếu người dùng nhập các ký tự đặt biệt vào các textbox
            {
                int id = int.Parse(txt_ID.Text);
                DateTime date = pk_Date.SelectedDate ?? DateTime.Now;   // Nếu người dùng không chọn ngày thì lấy ngày hiện tại
                string name = cb_name.Text;
                double price = double.Parse(txt_Price.Text);
                string moreInfo = txt_moreInfo.Text;
                TimeSpan time = TimeSpan.Parse(tp_time.Text);

                UpdateData(id, date, name, price, moreInfo, time);
                LoadCombobox();     //sau khi cập nhật tải lại cbbox
                ScrollDataGrid();   //sau khi cập nhật cuộn datagrid xuống
                TotalPrice();       //sau khi cập nhật sẽ tải lại tổng tiền
            }
            catch
            {
                MessageBox.Show("Có Lỗi Khi Cập Nhật Dữ Liệu\nVui Lòng Kiểm Tra Lại", "Thông Báo", MessageBoxButton.OK);
            }
        }

        private void bt_Del_Click(object sender, RoutedEventArgs e)     //Xoá dữ liệu
        {
            int id = int.Parse(txt_ID.Text);
            MessageBoxResult result = MessageBox.Show("Bạn có chắc muốn xoá không?", "Thông báo", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            if (result == MessageBoxResult.OK)
            {
                DeleteData(id);
                LoadCombobox();     //sau khi xoá load lại cbbox
                ScrollDataGrid();   //sau khi xoá cuộn datagird xuống
                TotalPrice();       //sau khi xoá sẻ tải lại tổng tiền
            }
            
        }

        private void txt_Search_TextChanged(object sender, TextChangedEventArgs e)  //Tìm Kiếm dữ liệu
        {
            SearchDataText();
        }

        private void dp_search_SelectedDateChanged(object sender, SelectionChangedEventArgs e)  //Tìm kiếm theo ngày
        {
            SearchDateSeleted();
        }

        private void bt_clearData_Click(object sender, RoutedEventArgs e)   //Dọn dữ liệu
        {
            ClearData();    //gọi hàm xoá dữ liệu
        }

        private void bt_Static_Click(object sender, RoutedEventArgs e)      //Thống kê
        {
            Statics wd = new Statics();
            wd.ShowDialog();
        }

        private void bt_excel_Click(object sender, RoutedEventArgs e)       //Xuất file excel
        {
            ExportToExcel();
        }

        private void bt_pdf_Click(object sender, RoutedEventArgs e)     //Xuất file pdf
        {
            ExportToPdf(dg_show);
        }

        private void dg_show_SelectionChanged(object sender, SelectionChangedEventArgs e)       //Binding dữ liệu lên textbox
        {
            CountDataItem();
            CalculateTotal();   //Gọi hàm tính tổng tiền
            QLCT qlct = (QLCT)dg_show.SelectedItem;
            if (qlct != null)
            {
                txt_ID.Text = qlct.Id.ToString();
                pk_Date.SelectedDate = (DateTime)qlct.Date;
                cb_name.Text = qlct.Name.ToString();
                txt_Price.Text = qlct.Price.ToString();
                txt_moreInfo.Text = qlct.MoreInfo.ToString();
            }
        }

        private void lsb_logout_PreviewMouseDown(object sender, MouseButtonEventArgs e)     //Menu đăng xuất
        {
            Application.Current.Shutdown();
        }

        private void lsb_account_PreviewMouseDown(object sender, MouseButtonEventArgs e)        //menu tài khoản
        {

        }

        private void lsb_setting_PreviewMouseDown(object sender, MouseButtonEventArgs e)        //menu cài đặt
        {

        }

        #endregion

        private void TEST_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
