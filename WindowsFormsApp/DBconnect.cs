using Microsoft.Data.SqlClient; // Sử dụng namespace Microsoft.Data.SqlClient để làm việc với SQL Server
using System.Data; // Sử dụng namespace System.Data để làm việc với các đối tượng dữ liệu
using System;
using System.IO;
using System.Windows.Forms;
namespace WindowsFormsApp
{
    internal class ConfigReader
    {
        // Phương thức để đọc các giá trị từ tệp cấu hình
        public static ConfigValues ReadConfig(string filePath)
        {
            ConfigValues configValues = new ConfigValues();
            try
            {
                // Đảm bảo tệp tồn tại
                if (File.Exists(filePath))
                {
                    // Đọc tất cả các dòng từ tệp
                    string[] lines = File.ReadAllLines(filePath);
                    foreach (string line in lines)
                    {
                        // Kiểm tra xem dòng có bắt đầu bằng "IP", "Database", "User" hoặc "Pass" không
                        if (line.StartsWith("IP="))
                        {
                            configValues.Server = line.Substring(3); // Lấy phần sau dấu "=" là giá trị của Server
                        }
                        else if (line.StartsWith("Database="))
                        {
                            configValues.Database = line.Substring(9); // Lấy phần sau dấu "=" là giá trị của Database
                        }
                        else if (line.StartsWith("User="))
                        {
                            configValues.UserId = line.Substring(5); // Lấy phần sau dấu "=" là giá trị của User ID
                        }
                        else if (line.StartsWith("Pass="))
                        {
                            configValues.Password = line.Substring(5); // Lấy phần sau dấu "=" là giá trị của Password
                        }
                    }
                }
                else
                {
                    Console.WriteLine("File not found: " + filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error reading config file: " + ex.Message);
            }
            return configValues;
        }
    }

    // Class để lưu trữ các giá trị cấu hình
    internal class ConfigValues
    {
        public string Server { get; set; }
        public string Database { get; set; }
        public string UserId { get; set; }
        public string Password { get; set; }
    }
    internal class DBconnect
    {
        // Chuỗi kết nối đến cơ sở dữ liệu SQL Server
        // Phương thức getter để lấy chuỗi kết nối
        // Phương thức để giải mã password từ dạng mã hóa
        // Phương thức để giải mã password từ dạng mã hóa
        // Phương thức để giải mã password từ dạng mã hóa
        private string DecryptPassword(string encryptedPassword)
        {
            string decryptedPassword = "";
            int ad_LT = int.Parse(encryptedPassword.Substring(0, 1));

            for (int i = 1; i < encryptedPassword.Length; i++)
            {
                char tempChar = encryptedPassword[i];
                int ASCII_int = Convert.ToInt32(tempChar) - ad_LT;
                decryptedPassword += Convert.ToString((char)ASCII_int);
            }

            return decryptedPassword;
        }



        // Chuỗi kết nối đến cơ sở dữ liệu SQL Server
        // Phương thức getter để lấy chuỗi kết nối
        // Chuỗi kết nối đến cơ sở dữ liệu SQL Server
        // Phương thức getter để lấy chuỗi kết nối
        public string connectionString
        {
            get
            {
                // Thay thế giá trị chuỗi kết nối bằng các giá trị từ tệp cấu hình, và giải mã password
                ConfigValues config = ConfigReader.ReadConfig(@"C:\ERP\ComName2.ini");
                string decryptedPassword = DecryptPassword(config.Password);
                return $"Server={config.Server};Database={config.Database};User Id={config.UserId};Password={decryptedPassword};TrustServerCertificate=true;";
            }
        }


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
        public int ExecuteSqlCommand(SqlCommand cmd)
        {
            using (SqlConnection connection = new SqlConnection(connectionString)) // Sử dụng SqlConnection để kết nối đến cơ sở dữ liệu
            {
                try
                {
                    cmd.Connection = connection;
                    connection.Open();
                    return cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    // Hiển thị thông báo lỗi
                    MessageBox.Show("Đã xảy ra lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // Trả về -1 hoặc một giá trị thích hợp để biểu thị lỗi
                    return -1;
                }
            }
        }

        //
        // Phương thức ExecuteScalar thực hiện truy vấn SQL và trả về một giá trị duy nhất
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
        public int ExecuteScalarInt(string query, SqlParameter[] parameters = null)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                if (parameters != null)
                {
                    command.Parameters.AddRange(parameters);
                }
                connection.Open();
                object result = command.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    int parsedResult;
                    if (int.TryParse(result.ToString(), out parsedResult))
                    {
                        return parsedResult;
                    }
                    else
                    {
                        throw new FormatException("Kết quả trả về không thể chuyển đổi sang kiểu int.");
                    }
                }
                else
                {
                    return 0;
                }
            }
        }

        //
    }

}
