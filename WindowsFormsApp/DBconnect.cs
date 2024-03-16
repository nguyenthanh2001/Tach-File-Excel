using Microsoft.Data.SqlClient; // Sử dụng namespace Microsoft.Data.SqlClient để làm việc với SQL Server
using System.Data; // Sử dụng namespace System.Data để làm việc với các đối tượng dữ liệu
using System;
namespace WindowsFormsApp
{
    internal class DBconnect
    {
        // Chuỗi kết nối đến cơ sở dữ liệu SQL Server
        string connectionString = "Server=192.168.60.60;Database=LIY_ERP;User Id=lacty;Password=wu0g3tp6;TrustServerCertificate=true;";
        //string connectionString = "Server=MSI;Database=laptop;User Id=sa;Password=123456;TrustServerCertificate=true;";

        // Phương thức ExecuteQuery thực hiện truy vấn SQL và trả về một DataTable chứa kết quả
        public DataTable ExecuteQuery(string query, SqlParameter[] parameters = null)
        {
            DataTable dataTable = new DataTable(); // Tạo một DataTable để lưu trữ kết quả truy vấn
            using (SqlConnection connection = new SqlConnection(connectionString)) // Sử dụng SqlConnection để kết nối đến cơ sở dữ liệu
            {
                SqlCommand command = new SqlCommand(query, connection); // Tạo một SqlCommand để thực thi truy vấn
                if (parameters != null) // Kiểm tra xem có tham số được truyền vào không
                {
                    command.Parameters.AddRange(parameters); // Thêm các tham số vào SqlCommand
                }
                SqlDataAdapter adapter = new SqlDataAdapter(command); // Tạo một SqlDataAdapter để lấp đầy dữ liệu từ cơ sở dữ liệu vào DataTable
                adapter.Fill(dataTable); // Lấp đầy DataTable với dữ liệu từ truy vấn
            }
            return dataTable; // Trả về DataTable chứa kết quả truy vấn
        }

        // Phương thức ExecuteNonQuery thực hiện truy vấn SQL và trả về số hàng bị ảnh hưởng
        public int ExecuteNonQuery(string query, SqlParameter[] parameters = null)
        {
            using (SqlConnection connection = new SqlConnection(connectionString)) // Sử dụng SqlConnection để kết nối đến cơ sở dữ liệu
            {
                SqlCommand command = new SqlCommand(query, connection); // Tạo một SqlCommand để thực thi truy vấn
                if (parameters != null) // Kiểm tra xem có tham số được truyền vào không
                {
                    command.Parameters.AddRange(parameters); // Thêm các tham số vào SqlCommand
                }
                connection.Open(); // Mở kết nối đến cơ sở dữ liệu
                return command.ExecuteNonQuery(); // Thực thi truy vấn và trả về số hàng bị ảnh hưởng
            }
        }
        //
        // Phương thức ExecuteScalar thực hiện truy vấn SQL và trả về một giá trị duy nhất
        public int ExecuteScalar(string query, SqlParameter[] parameters = null)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                if (parameters != null)
                {
                    command.Parameters.AddRange(parameters);
                }
                connection.Open();
                // Thực hiện truy vấn và trả về giá trị đầu tiên của hàng đầu tiên trong tập kết quả
                object result = command.ExecuteScalar();
                // Kiểm tra xem giá trị có tồn tại không trước khi chuyển đổi sang kiểu int
                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToInt32(result);
                }
                else
                {
                    // Trả về 0 nếu không có giá trị trả về
                    return 0;
                }
            }
        }
        //
    }

}
