using System;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;

namespace WindowsFormsApp
{
    public partial class LoginForm : Form
    {
        private DBconnect dbConnect;

        // Định nghĩa một thuộc tính để lưu giá trị username
        public string Username { get; private set; }

        public LoginForm()
        {
            InitializeComponent();
            dbConnect = new DBconnect(); // Khởi tạo đối tượng DBconnect
                                         // Đặt vị trí của form ở giữa màn hình
            this.StartPosition = FormStartPosition.CenterScreen;

            // Ẩn mật khẩu khi khởi động ứng dụng
            tb_password.UseSystemPasswordChar = true;

            //gan ten nguoi dung gan nhat
            if (System.IO.File.Exists("lastusername_xoaytua.txt"))
            {
                string lastUsername = System.IO.File.ReadAllText("lastusername_xoaytua.txt");
                // Điền tên người dùng gần nhất vào TextBox tương ứng
                tb_account.Text = lastUsername;
            }
            if (System.IO.File.Exists("lastpassword_xoaytua.txt"))
            {
                string lastPassWord = System.IO.File.ReadAllText("lastpassword_xoaytua.txt");
                string decryptedPassword = PasswordEncryption.DecryptPassword(lastPassWord);
                // Điền mật khẩu giải mã vào TextBox tương ứng
                tb_password.Text = decryptedPassword;
            }
        }

        private void btn_login_Click(object sender, EventArgs e)
        {
            string username = tb_account.Text;
            string password = tb_password.Text;
            string encryptedPassword = PasswordEncryption.EncryptPassword(password);
            // Kiểm tra nếu cả hai biến đều rỗng
            if (string.IsNullOrEmpty(username) && string.IsNullOrEmpty(password))
            {
                // Hiển thị thông báo
                MessageBox.Show("Vui lòng không bỏ trống tài khoản và mật khẩu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // Dừng lại và không thực hiện bất kỳ hành động nào khác
                return;
            }
            // Thực hiện truy vấn kiểm tra thông tin đăng nhập với cơ sở dữ liệu
            string query = "SELECT COUNT(*) FROM Busers WHERE USERID = @username AND PWD = @password";
            SqlParameter[] parameters =
            {
                new SqlParameter("@username", System.Data.SqlDbType.VarChar) { Value = username },
                new SqlParameter("@password", System.Data.SqlDbType.VarChar) { Value = password }
            };

            int userCount = dbConnect.ExecuteScalar(query, parameters);

            if (userCount > 0)
            {
                // Lưu giá trị của biến username
                Username = username;
                // Lưu tên người dùng vào tệp văn bản
                System.IO.File.WriteAllText("lastusername_xoaytua.txt", username);

                // Mã hóa mật khẩu trước khi lưu vào cơ sở dữ liệu
                
                System.IO.File.WriteAllText("lastpassword_xoaytua.txt", encryptedPassword);

                // Đăng nhập thành công
                this.DialogResult = DialogResult.OK;
                //
                // Đăng nhập thành công, ẩn LoginForm
                //this.Hide();
                MessageBox.Show("Đăng nhập thành công! Đang xử lý...", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // Đăng nhập không thành công
                MessageBox.Show("Tên người dùng hoặc mật khẩu không chính xác!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //
        private void checkBoxShowPassword_CheckedChanged(object sender, EventArgs e)
        {
            // Nếu checkbox được kiểm tra, hiển thị mật khẩu
            // Nếu không, ẩn mật khẩu
            tb_password.UseSystemPasswordChar = !CB_Showpass.Checked;
        }
       
        //
        private void tb_account_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Down)
            {
                tb_password.Focus();
                e.SuppressKeyPress = true; // Ngăn không cho kí tự Enter được thêm vào TextBox1
            }
        }
        private void TB_password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Down)
            {
                btn_login.Focus();
                e.SuppressKeyPress = true; // Ngăn không cho kí tự Enter được thêm vào TextBox1
            }
        }
        //ma hoa mat khau
        

        //

    }
    public static class PasswordEncryption
    {
        // Ma hoa
        public static string EncryptPassword(string password)
        {
            // Thêm một bước trộn chuỗi
            string mixedPassword = MixString(password);

            // Chia nhỏ chuỗi và đảo ngược từng phần
            string[] parts = SplitString(mixedPassword);
            for (int i = 0; i < parts.Length; i++)
            {
                char[] charArray = parts[i].ToCharArray();
                Array.Reverse(charArray);
                parts[i] = new string(charArray);
            }

            // Ghép các phần đã biến đổi lại thành một chuỗi kết quả
            return string.Join("", parts);
        }

        // Giai ma
        public static string DecryptPassword(string encryptedPassword)
        {
            // Chia nhỏ chuỗi và đảo ngược từng phần
            string[] parts = SplitString(encryptedPassword);
            for (int i = 0; i < parts.Length; i++)
            {
                char[] charArray = parts[i].ToCharArray();
                Array.Reverse(charArray);
                parts[i] = new string(charArray);
            }

            // Ghép các phần đã biến đổi lại thành một chuỗi kết quả
            string mixedPassword = string.Join("", parts);

            // Loại bỏ bước trộn để giải mã
            return UnmixString(mixedPassword);
        }

        // Hàm trộn chuỗi: hoán đổi ký tự đầu và cuối chuỗi và thêm một ký tự đặc biệt vào giữa
        private static string MixString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            char[] charArray = input.ToCharArray();
            if (charArray.Length >= 2)
            {
                char temp = charArray[0];
                charArray[0] = charArray[charArray.Length - 1];
                charArray[charArray.Length - 1] = temp;
            }
            string mixedString = new string(charArray);

            // Thêm một ký tự đặc biệt vào giữa chuỗi
            return mixedString.Insert(mixedString.Length / 2, "#");
        }

        // Hàm loại bỏ bước trộn chuỗi: loại bỏ ký tự đặc biệt và hoán đổi lại ký tự đầu và cuối
        private static string UnmixString(string mixedInput)
        {
            // Loại bỏ ký tự đặc biệt
            string unmixString = mixedInput.Replace("#", "");

            // Hoán đổi lại ký tự đầu và cuối chuỗi
            char[] charArray = unmixString.ToCharArray();
            if (charArray.Length >= 2)
            {
                char temp = charArray[0];
                charArray[0] = charArray[charArray.Length - 1];
                charArray[charArray.Length - 1] = temp;
            }
            return new string(charArray);
        }

        // Hàm chia nhỏ chuỗi: chia chuỗi thành các phần có độ dài bằng nhau
        private static string[] SplitString(string input)
        {
            // Chia chuỗi thành các phần có độ dài bằng nhau (nếu độ dài của chuỗi không chia hết cho 4, có thể có phần dư)
            int partLength = input.Length / 4;
            int remainingLength = input.Length % 4;
            string[] parts = new string[4];
            int startIndex = 0;
            for (int i = 0; i < 4; i++)
            {
                int length = partLength;
                if (i < remainingLength)
                    length++;

                parts[i] = input.Substring(startIndex, length);
                startIndex += length;
            }
            return parts;
        }

    }
}
