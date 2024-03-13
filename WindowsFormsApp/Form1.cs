using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Data;
using System.Drawing;

namespace WindowsFormsApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Text = "Tách sheet từ tệp Excel";
            // Hiệu ứng

        }
        // Hiệu ứng
        // Hiệu ứng

        // Hiệu ứng
        // Hiệu ứng
        //
        /*
        private void button1_Click(object sender, EventArgs e)
        {
            // Mặc định đường dẫn vào thư mục C:\ và thư mục ERP và tên tệp Excel "filexoaytua_input.xlsx"
            string defaultPath = @"C:\ERP\filexoaytua_input.xlsx";

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
        */

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
                        // Cập nhật trạng thái xử lý
                        UpdateStatusLabel($"Processing sheet '{worksheet.Name}', row {row}...");

                        // Kiểm tra điều kiện ban đầu (gặp ký tự 主管)
                        if (worksheet.Cells[row, 5].Text.ToLower() == "主管﹕")
                        {
                            newConditionMet = true; // Đặt biến để báo hiệu điều kiện mới đã được đáp ứng
                                                    // Tạo một worksheet mới trong workbook mới
                            ExcelWorksheet newWorksheet = newPackage.Workbook.Worksheets.Add($"{worksheet.Name}_Page{lastRowOfPage}");

                            // Copy dữ liệu từ trang hiện tại sang trang mới
                            /*
                            for (int i = lastRowOfPage; i <= row; i++)
                            {
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    newWorksheet.Cells[i - lastRowOfPage + 1, col].Value = worksheet.Cells[i, col].Value;
                                }
                            }
                            */
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


                // Hiển thị số lượng sheet đã tạo mới trên Label
                UpdateSheetCountLabel(numberOfSheetsCreated);
                // Yêu cầu người dùng chọn thư mục để lưu file Excel mới

                // Lấy đường dẫn và tên file input
                string inputFileName = Path.GetFileNameWithoutExtension(inputFile);
                string inputDirectory = Path.GetDirectoryName(inputFile);

                // Tạo tên file mới và đường dẫn đến thư mục input
                string outputFileName = Path.Combine(inputDirectory, "filexoaytua_output.xlsx");

                // Lưu workbook mới vào cùng thư mục với file input
                newPackage.SaveAs(new FileInfo(outputFileName));

                // Hiển thị thông báo thành công
                MessageBox.Show("Completed");
                /*
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.RestoreDirectory = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Lưu workbook mới vào thư mục được chọn
                    string outputFileName = saveFileDialog.FileName;
                    newPackage.SaveAs(new FileInfo(outputFileName));

                    // Hiển thị thông báo thành công
                    MessageBox.Show("Tách sheet thành công và lưu vào: " + outputFileName, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Hiển thị thông báo nếu người dùng không chọn thư mục
                    MessageBox.Show("Chưa chọn thư mục lưu file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                */
                //
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
            string defaultPath = @"C:\ERP\filexoaytua_input.xlsx";

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

