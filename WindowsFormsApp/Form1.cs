using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Text.RegularExpressions;

namespace WindowsFormsApp
{
    public partial class Form1 : Form
    {
        private DBconnect dbConnect;
        public Form1()
        {
            InitializeComponent();
            dbConnect = new DBconnect(); // Khởi tạo kết nối cơ sở dữ liệu
            this.Text = "Tách sheet từ tệp Excel";
            // Hiệu ứng

        }
        //

        private void ProcessExcelFile(string inputFile)
        {

            // Thiết lập LicenseContext cho EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Đọc tệp Excel đầu vào
            FileInfo fileInfo = new FileInfo(inputFile);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorkbook workbook = package.Workbook;

                // Tạo một workbook mới để lưu các sheet mới
                ExcelPackage newPackage = new ExcelPackage();

                // Biến đếm số lượng sheet đã tạo mới
                int numberOfSheetsCreated = 0;

                // Tạo DataTable để lưu trữ dữ liệu từ tệp Excel
                DataTable excelData = new DataTable();
                excelData.Columns.Add("ProNo"); // Thêm cột ProNo vào DataTable để lưu giá trị này từ Excel
                                                // Thêm các cột tương ứng với dữ liệu bạn muốn lưu trữ từ tệp Excel
                excelData.Columns.Add("Lean");
                excelData.Columns.Add("Ten_Giay");
                excelData.Columns.Add("Dao_Chat");
                excelData.Columns.Add("Article");
                excelData.Columns.Add("Dang_Fom");
                excelData.Columns.Add("Goo");
                excelData.Columns.Add("May");
                excelData.Columns.Add("Chat");
                excelData.Columns.Add("Ry");
                excelData.Columns.Add("Size");

                // Lặp qua từng sheet trong workbook 
                foreach (ExcelWorksheet worksheet in workbook.Worksheets)
                {
                    // Khởi tạo biến để lưu vị trí hàng cuối cùng của trang
                    int lastRowOfPage = 1;

                    // Biến để kiểm tra điều kiện mới
                    bool newConditionMet = false;

                    // Lặp qua từng hàng trong sheet
                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        // Tạo một hàng mới trong DataTable excelData
                        DataRow newRow = excelData.NewRow();

                        // Đọc giá trị từ cột A (ProNo) và gán vào cột ProNo của DataRow
                        newRow["ProNo"] = "111111111";

                        string pattern1 = ".*ĐÓNG.*";
                        object cellValue = worksheet.Cells[row, 9].Value;
                        if (cellValue != null && !cellValue.ToString().Equals("訂單號碼 \nRY") && !cellValue.ToString().Equals("ĐÓNG ĐƠN 10**/TH") && !Regex.IsMatch(cellValue.ToString(), pattern1)) 
                        {
                            newRow["Lean"] = cellValue.ToString();
                        }
                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Lean"] = string.Empty; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //
                        object cellValue1 = worksheet.Cells[row, 2].Value;
                        if (cellValue1 != null && !cellValue1.ToString().Equals("型體名稱\nTÊN GIÀY") && !cellValue1.ToString().Equals("合計：") && !cellValue1.ToString().Equals("型\n預\n計\n生\n產\n時\n間")) 
                        {
                            newRow["Ten_Giay"] = worksheet.Cells[row, 2].Value.ToString(); 

                        }
                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Ten_Giay"] = string.Empty;
                        }
                        //
                        string patterndc = ".*DAO CHẶT.*";
                        object cellValuedc = worksheet.Cells[row, 3].Value;
                        if (cellValuedc != null && !Regex.IsMatch(cellValuedc.ToString(), patterndc))
                        {
                            newRow["Dao_Chat"] = cellValue.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Dao_Chat"] = string.Empty; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //
                        string patternart = ".*ARTICLE.*";
                        object cellValueart = worksheet.Cells[row, 4].Value;
                        if (cellValueart != null && !Regex.IsMatch(cellValueart.ToString(), patternart))
                        {
                            newRow["Article"] = cellValueart.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Article"] = string.Empty; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //Dang_Fom
                        string patterndf = ".*DẠNG FOM.*";
                        object cellValuedf = worksheet.Cells[row, 5].Value;
                        if (cellValuedf != null && !Regex.IsMatch(cellValuedf.ToString(), patterndf))
                        {
                            newRow["Dang_Fom"] = cellValuedf.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Dang_Fom"] = string.Empty; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //Goo
                        string patterg = ".*GÒ.*";
                        object cellValueg = worksheet.Cells[row, 6].Value;
                        if (cellValueg != null && !Regex.IsMatch(cellValueg.ToString(), patterg))
                        {
                            newRow["Dang_Fom"] = cellValueg.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Dang_Fom"] = string.Empty; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        // Thêm DataRow mới vào DataTable
                        excelData.Rows.Add(newRow); // Không cần phải sử dụng Clone() ở đây
                        //May
                        string pattenm = ".*MAY.*";
                        object cellValuem = worksheet.Cells[row, 7].Value;
                        if (cellValuem != null && !Regex.IsMatch(cellValuem.ToString(), pattenm))
                        {
                            newRow["May"] = cellValuem.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["May"] = string.Empty; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //Chat
                        string pattenchat = ".*MAY.*";
                        object cellValuechat = worksheet.Cells[row, 8].Value;
                        if (cellValuechat != null && !Regex.IsMatch(cellValuechat.ToString(), pattenchat))
                        {
                            newRow["Chat"] = cellValuechat.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Chat"] = string.Empty; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }


                        // Cập nhật trạng thái xử lý May Chat Ry Size
                        UpdateStatusLabel($"Processing sheet '{worksheet.Name}', row {row}...");

                        // Kiểm tra điều kiện ban đầu (gặp ký tự 主管)
                        if (worksheet.Cells[row, 5].Text.ToLower() == "主管﹕")
                        {
                            newConditionMet = true; // Đặt biến để báo hiệu điều kiện mới đã được đáp ứng
                                                    // Tạo một worksheet mới trong workbook mới
                            ExcelWorksheet newWorksheet = newPackage.Workbook.Worksheets.Add($"{worksheet.Name}_Page{lastRowOfPage}");
                            // Copy dữ liệu từ trang hiện tại sang trang mới
                            for (int i = lastRowOfPage; i <= row; i++)
                            {
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    if (worksheet.Cells[i, col].Text == "#REF!")
                                    {
                                        newWorksheet.Cells[i - lastRowOfPage + 1, col].Value = 0;
                                    }
                                    else
                                    {
                                        newWorksheet.Cells[i - lastRowOfPage + 1, col].Value = worksheet.Cells[i, col].Value;
                                    }

                                }
                            }
                            // Tăng biến đếm số lượng sheet đã tạo mới
                            numberOfSheetsCreated++;
                            //
                            // Cập nhật vị trí hàng cuối cùng của trang
                            lastRowOfPage = row + 1;
                        }
                        // Kiểm tra điều kiện mới (ví dụ: gặp ký tự khác)
                        else if (worksheet.Cells[row, 1].Text.ToLower() == "bang xoay tua  ( may )")
                        {
                            // Nếu điều kiện mới được đáp ứng và điều kiện ban đầu cũng đã được đáp ứng
                            // thì tạo một trang mới
                            if (newConditionMet)
                            {
                                // Tạo một worksheet mới trong workbook mới
                                ExcelWorksheet newWorksheet = worksheet;
                                // Copy dữ liệu từ trang hiện tại sang trang mới
                                for (int i = lastRowOfPage; i <= row; i++)
                                {
                                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                    {
                                        // Sao chép dữ liệu tại sheet hiện tại
                                        worksheet.Cells[i, col].Value = worksheet.Cells[i, col].Value;
                                    }
                                }

                                // Cập nhật vị trí hàng cuối cùng của trang
                                lastRowOfPage = row;
                            }
                        }                      
                    }
                    // Cập nhật lastRowOfPage tại đây
                    lastRowOfPage = 1;
                }

                // Thêm hàng mới vào DataTable
                SaveDataToDatabase(excelData);


                // Hiển thị số lượng sheet đã tạo mới trên Label
                UpdateSheetCountLabel(numberOfSheetsCreated);


                // Hiển thị thông báo thành công
                MessageBox.Show("Completed");
                //
            }
        }

        private void SaveDataToDatabase(DataTable data)
        {
            // Chuẩn bị truy vấn SQL để chèn dữ liệu vào cơ sở dữ liệu
            string query = "INSERT INTO BANG_XOAY_TUA (ProNo, Lean, Ten_Giay, Dao_Chat, Article, Dang_Fom, Goo, May, Chat, Ry, Size) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5, @Value6, @Value7, @Value8, @Value9, @Value10, @Value11)";

            // Lặp qua từng hàng trong DataTable và chèn dữ liệu vào cơ sở dữ liệu
            foreach (DataRow row in data.Rows)
            {
                // Tạo mảng tham số để truyền giá trị vào truy vấn SQL Article
                SqlParameter[] parameters =
                {
            new SqlParameter("@Value1", SqlDbType.VarChar) { Value = row["ProNo"] },
            new SqlParameter("@Value2", SqlDbType.VarChar) { Value = row["Lean"] },
            new SqlParameter("@Value3", SqlDbType.VarChar) { Value = row["Ten_Giay"] },
            new SqlParameter("@Value4", SqlDbType.VarChar) { Value = row["Dao_Chat"] },
            new SqlParameter("@Value5", SqlDbType.VarChar) { Value = row["Article"] },
            new SqlParameter("@Value6", SqlDbType.VarChar) { Value = row["Dang_Fom"] },
            new SqlParameter("@Value7", SqlDbType.VarChar) { Value = row["Goo"] },
            new SqlParameter("@Value8", SqlDbType.VarChar) { Value = row["May"] },
            new SqlParameter("@Value9", SqlDbType.VarChar) { Value = row["Chat"] },
            new SqlParameter("@Value10", SqlDbType.VarChar) { Value = row["Ry"] },
            new SqlParameter("@Value11", SqlDbType.VarChar) { Value = row["Size"] }
            

            
            // Thêm các tham số cho các cột khác nếu cần thiết 
                };

                // Thực thi truy vấn chèn dữ liệu vào cơ sở dữ liệu
                dbConnect.ExecuteNonQuery(query, parameters);
            }
        }
        private void UpdateStatusLabel(string message)
        {
            // Đảm bảo gọi cập nhật trên luồng UI chính
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { UpdateStatusLabel(message); }));
                return;
            }
            // Cập nhật nhãn trạng thái với thông điệp mới
            label1.Text = message;
            // Cập nhật trạng thái về giao diện người dùng
            Application.DoEvents();
        }
        private void UpdateSheetCountLabel(int numberOfSheets)
        {
            // Cập nhật nhãn Label để hiển thị số lượng sheet đã tạo mới
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate { UpdateSheetCountLabel(numberOfSheets); }));
                return;
            }
            label1.Text = $"Tổng số sheet đã tạo mới: {numberOfSheets}";
            Application.DoEvents();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // Mặc định đường dẫn vào thư mục C:\ và thư mục ERP và tên tệp Excel "filexoaytua_input.xlsx"
            string defaultPath = @"C:\ERPP\filexoaytua_input.xlsx";

            // Kiểm tra sự tồn tại của tệp Excel
            if (File.Exists(defaultPath))
            {
                // Nếu tệp tồn tại, tiến hành xử lý
                ProcessExcelFile(defaultPath);
            }
            else
            {
                // Nếu không tìm thấy tệp, hiển thị thông báo
                MessageBox.Show("Không tìm thấy tệp Excel filexoaytua_input.xlsx trong thư mục C:\\ERP\\filexoaytua_input", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //
    }
}

