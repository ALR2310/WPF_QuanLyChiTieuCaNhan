using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuanLyChiTieu.Model
{
    public class DataProvider
    {
        private static DataProvider instance;
        public static DataProvider Instance
        {
            get { if (instance == null) instance = new DataProvider(); return DataProvider.instance; }
            private set { DataProvider.instance = value; }
        }

        private DataProvider() { }

        private string connectSTRING = "Data Source=ANRIP\\SQLEXPRESS;Initial Catalog=QuanLyChiTieu;Integrated Security=True";

        public DataTable ExecuteQuery(string query)
        {
            using (SqlConnection connection = new SqlConnection(connectSTRING))
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand(query, connection);
                DataTable data = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(data);
                connection.Close();
                return data;
            }
        }

        public int ExecuteNonQuery(string query, object[] parameter = null)
        {
            int data = 0;
            using (SqlConnection ketnoi = new SqlConnection(connectSTRING))
            {
                ketnoi.Open();
                SqlCommand cmd = new SqlCommand(query, ketnoi);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            cmd.Parameters.AddWithValue(item, parameter[i]);
                            i++;
                        }
                    }
                }
                data = cmd.ExecuteNonQuery();
                ketnoi.Close();
                return data;
            }
        }
    }
}
