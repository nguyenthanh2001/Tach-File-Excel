using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Security.Principal;
using System.Globalization;

namespace WindowsFormsApp
{
    public partial class Form1 : Form
    {
        private Timer timer;
        private DBconnect dbConnect;
        private long ProNo; // Bắt đầu từ 0
        public Form1()
        {
            InitializeComponent();
            dbConnect = new DBconnect(); // Khởi tạo kết nối cơ sở dữ liệu

            this.Text = "Tách sheet từ tệp Excel";

        }
        //
        public int GetNextIdentityValueFromDatabase()
        {
            // Chuỗi truy vấn SQL để lấy 5 số cuối cùng của giá trị prono
            string query = "SELECT RIGHT(ISNULL(MAX(prono), '00000'), 5) FROM ProgressPross";

            // T
            // Thực hiện truy vấn và trả về giá trị lớn nhất hiện có của cột tự tăng
            return dbConnect.ExecuteScalar(query);
        }

        //
        private void ProcessExcelFile(string inputFile)
        {
            // Lấy giá trị năm và tháng từ thời gian hiện tại
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;


            // Lấy giá trị tự tăng từ cơ sở dữ liệu
            object idsqlObject = GetNextIdentityValueFromDatabase();
            long idsql = 0;

            if (idsqlObject != null && idsqlObject != DBNull.Value)
            {
                idsql = Convert.ToInt64(idsqlObject);
            }
            else
            {
                idsql = 0;
            }
            ProNo = idsql+1;
            // Tiếp tục xử lý với giá trị idsql đã kiểm tra




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
                //int numberOfSheetsCreated = 0;

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
                excelData.Columns.Add("Seq1");
                excelData.Columns.Add("Seq2");
                excelData.Columns.Add("Seq3");
                excelData.Columns.Add("Seq4");
                excelData.Columns.Add("Seq5");
                excelData.Columns.Add("Seq6");
                excelData.Columns.Add("Seq7");
                excelData.Columns.Add("Seq8");
                excelData.Columns.Add("Seq9");
                excelData.Columns.Add("Seq10");
                excelData.Columns.Add("Seq11");
                excelData.Columns.Add("Seq12");
                excelData.Columns.Add("Seq13");
                excelData.Columns.Add("Seq14");
                excelData.Columns.Add("Seq15");
                excelData.Columns.Add("Seq16");
                excelData.Columns.Add("Seq17");
                excelData.Columns.Add("Seq18");
                excelData.Columns.Add("Seq19");
                excelData.Columns.Add("Seq20");
                excelData.Columns.Add("Seq21");
                excelData.Columns.Add("Seq22");
                excelData.Columns.Add("Seq23");
                excelData.Columns.Add("Seq24");
                excelData.Columns.Add("Seq25");
                excelData.Columns.Add("SO_CHI_LENH");
                excelData.Columns.Add("THUC_TE_PC");
                excelData.Columns.Add("LUY_TICH_PC");
                excelData.Columns.Add("SO_CHUA_PC");
                excelData.Columns.Add("Don_Vi_San_Xuat");
                excelData.Columns.Add("Created_Date");
                
                //
                //
                // Lấy giá trị năm và tháng từ thời gian hiện tại

                // Lặp qua từng sheet trong workbook 
                foreach (ExcelWorksheet worksheet in workbook.Worksheets)
                {

                    // Lặp qua từng hàng trong sheet
                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        // Tạo một hàng mới trong DataTable excelData
                        DataRow newRow = excelData.NewRow();

                        
                        //xu ly pro no
                        

                        // Gán giá trị ProNo cho newRow["ProNo"]
                        newRow["ProNo"] = ProNo;

                        //xu ly pro no
                        object cellValue1 = worksheet.Cells[row, 2].Value;
                        if (cellValue1 != null && !cellValue1.ToString().Equals("型體名稱\nTÊN GIÀY") && !cellValue1.ToString().Equals("合計：") && !cellValue1.ToString().Equals("型\n預\n計\n生\n產\n時\n間"))
                        {
                            newRow["Ten_Giay"] = worksheet.Cells[row, 2].Value.ToString();

                        }
                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Ten_Giay"] = null;
                        }
                        //
                        string patterndc = ".*DAO CHẶT.*";
                        object cellValuedc = worksheet.Cells[row, 3].Value;
                        if (cellValuedc != null && !Regex.IsMatch(cellValuedc.ToString(), patterndc))
                        {
                            newRow["Dao_Chat"] = cellValuedc.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Dao_Chat"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
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
                            newRow["Article"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //Dang_Fom
                        string patterndf = ".*DẠNG FOM.*";
                        object cellValuedf = worksheet.Cells[row, 5].Value;
                        if (cellValuedf != null && !Regex.IsMatch(cellValuedf.ToString(), patterndf))
                        {
                            newRow["Dang_Fom"] = cellValuedf.ToString();
                        }
                        //主管﹕ if (cellValue1 != null && !cellValue1.ToString().Equals("型體名稱\nTÊN GIÀY")
                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Dang_Fom"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }






                        if (cellValuedf != null && cellValuedf.ToString().Equals("主管﹕"))
                        {
                            //int identity = GetNextIdentityValueFromDatabase();
                            //ProNo = ProNo + identity;
                            //newRow["ProNo"] = ProNo;
                            ProNo++;
                        }
                        // Format ProNo theo quy tắc year + month + 5 số tự tăng
                        string paddedCounter = ProNo.ToString().PadLeft(5, '0');
                        string proNo = year.ToString() + month.ToString("00") + paddedCounter;

                        // Gán giá trị ProNo cho newRow["ProNo"]
                        newRow["ProNo"] = proNo;

                        // Thêm newRow vào DataTable excelData
                        excelData.Rows.Add(newRow);






                        //
                        //Goo
                        string patterg = ".*GÒ.*";
                        object cellValueg = worksheet.Cells[row, 6].Value;
                        if (cellValueg != null && !Regex.IsMatch(cellValueg.ToString(), patterg))
                        {
                            newRow["Goo"] = cellValueg.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Goo"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        // Thêm DataRow mới vào DataTable
                        //excelData.Rows.Add(newRow); // Không cần phải sử dụng Clone() ở đây
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
                            newRow["May"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //Chat
                        string pattenchat = ".*CHẶT.*";
                        object cellValuechat = worksheet.Cells[row, 8].Value;
                        if (cellValuechat != null && !Regex.IsMatch(cellValuechat.ToString(), pattenchat))
                        {
                            newRow["Chat"] = cellValuechat.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Chat"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //Ry
                        string combinedPattern = ".*訂單號碼.*|.*ĐÓNG ĐƠN.*|.*TỔNG HỢP.*";

                        object cellValueRy = worksheet.Cells[row, 9].Value;
                        if (cellValueRy != null && !Regex.IsMatch(cellValueRy.ToString(), combinedPattern))
                        {
                            newRow["Ry"] = cellValueRy.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Ry"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //Size
                        // Xử lý cho cột Seq1

                        object cellValueSeq1 = worksheet.Cells[row, 10].Value;
                        if (cellValueSeq1 != null)
                        {
                            if (decimal.TryParse(cellValueSeq1.ToString(), out decimal seq1Value))
                            {
                                newRow["Seq1"] = seq1Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq1"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq1"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq2 đến Seq25 tương tự như trên...
                        // Size 2
                        object cellValueSeq2 = worksheet.Cells[row, 11].Value;
                        if (cellValueSeq2 != null)
                        {
                            if (decimal.TryParse(cellValueSeq2.ToString(), out decimal seq2Value))
                            {
                                newRow["Seq2"] = seq2Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq2"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq2"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }
                        // Xử lý cho cột Seq3
                        object cellValueSeq3 = worksheet.Cells[row, 12].Value;
                        if (cellValueSeq3 != null)
                        {
                            if (decimal.TryParse(cellValueSeq3.ToString(), out decimal seq3Value))
                            {
                                newRow["Seq3"] = seq3Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq3"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq3"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq4 đến Seq25 tương tự như trên...
                        // Size 4
                        object cellValueSeq4 = worksheet.Cells[row, 13].Value;
                        if (cellValueSeq4 != null)
                        {
                            if (decimal.TryParse(cellValueSeq4.ToString(), out decimal seq4Value))
                            {
                                newRow["Seq4"] = seq4Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq4"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq4"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }
                        // Xử lý cho cột Seq5
                        object cellValueSeq5 = worksheet.Cells[row, 14].Value;
                        if (cellValueSeq5 != null)
                        {
                            if (decimal.TryParse(cellValueSeq5.ToString(), out decimal seq5Value))
                            {
                                newRow["Seq5"] = seq5Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq5"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq5"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq6 đến Seq25 tương tự như trên...
                        // Size 6
                        object cellValueSeq6 = worksheet.Cells[row, 15].Value;
                        if (cellValueSeq6 != null)
                        {
                            if (decimal.TryParse(cellValueSeq6.ToString(), out decimal seq6Value))
                            {
                                newRow["Seq6"] = seq6Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq6"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq6"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }
                        // Xử lý cho cột Seq7
                        object cellValueSeq7 = worksheet.Cells[row, 16].Value;
                        if (cellValueSeq7 != null)
                        {
                            if (decimal.TryParse(cellValueSeq7.ToString(), out decimal seq7Value))
                            {
                                newRow["Seq7"] = seq7Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq7"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq7"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq8 đến Seq25 tương tự như trên...
                        // Size 8
                        object cellValueSeq8 = worksheet.Cells[row, 17].Value;
                        if (cellValueSeq8 != null)
                        {
                            if (decimal.TryParse(cellValueSeq8.ToString(), out decimal seq8Value))
                            {
                                newRow["Seq8"] = seq8Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq8"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq8"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }
                        // Xử lý cho cột Seq9
                        object cellValueSeq9 = worksheet.Cells[row, 18].Value;
                        if (cellValueSeq9 != null)
                        {
                            if (decimal.TryParse(cellValueSeq9.ToString(), out decimal seq9Value))
                            {
                                newRow["Seq9"] = seq9Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq9"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq9"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq10 đến Seq25 tương tự như trên...
                        // Size 10
                        object cellValueSeq10 = worksheet.Cells[row, 19].Value;
                        if (cellValueSeq10 != null)
                        {
                            if (decimal.TryParse(cellValueSeq10.ToString(), out decimal seq10Value))
                            {
                                newRow["Seq10"] = seq10Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq10"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq10"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }
                        // Xử lý cho cột Seq11
                        object cellValueSeq11 = worksheet.Cells[row, 20].Value;
                        if (cellValueSeq11 != null)
                        {
                            if (decimal.TryParse(cellValueSeq11.ToString(), out decimal seq11Value))
                            {
                                newRow["Seq11"] = seq11Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq11"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq11"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq12 đến Seq25 tương tự như trên...
                        // Size 12
                        object cellValueSeq12 = worksheet.Cells[row, 21].Value;
                        if (cellValueSeq12 != null)
                        {
                            if (decimal.TryParse(cellValueSeq12.ToString(), out decimal seq12Value))
                            {
                                newRow["Seq12"] = seq12Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq12"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq12"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }
                        // Xử lý cho cột Seq13
                        object cellValueSeq13 = worksheet.Cells[row, 22].Value;
                        if (cellValueSeq13 != null)
                        {
                            if (decimal.TryParse(cellValueSeq13.ToString(), out decimal seq13Value))
                            {
                                newRow["Seq13"] = seq13Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq13"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq13"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq14 đến Seq25 tương tự như trên...
                        // Size 14
                        object cellValueSeq14 = worksheet.Cells[row, 23].Value;
                        if (cellValueSeq14 != null)
                        {
                            if (decimal.TryParse(cellValueSeq14.ToString(), out decimal seq14Value))
                            {
                                newRow["Seq14"] = seq14Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq14"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq14"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }
                        // Xử lý cho cột Seq15
                        object cellValueSeq15 = worksheet.Cells[row, 24].Value;
                        if (cellValueSeq15 != null)
                        {
                            if (decimal.TryParse(cellValueSeq15.ToString(), out decimal seq15Value))
                            {
                                newRow["Seq15"] = seq15Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq15"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq15"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq16 đến Seq25 tương tự như trên...
                        // Size 16
                        object cellValueSeq16 = worksheet.Cells[row, 25].Value;
                        if (cellValueSeq16 != null)
                        {
                            if (decimal.TryParse(cellValueSeq16.ToString(), out decimal seq16Value))
                            {
                                newRow["Seq16"] = seq16Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq16"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq16"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq17 đến Seq25 tương tự như trên...
                        // Và tiếp tục cho tất cả các cột Seq tương ứng...
                        // Xử lý cho cột Seq17
                        object cellValueSeq17 = worksheet.Cells[row, 26].Value;
                        if (cellValueSeq17 != null)
                        {
                            if (decimal.TryParse(cellValueSeq17.ToString(), out decimal seq17Value))
                            {
                                newRow["Seq17"] = seq17Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq17"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq17"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq18 đến Seq25 tương tự như trên...
                        // Size 18
                        object cellValueSeq18 = worksheet.Cells[row, 27].Value;
                        if (cellValueSeq18 != null)
                        {
                            if (decimal.TryParse(cellValueSeq18.ToString(), out decimal seq18Value))
                            {
                                newRow["Seq18"] = seq18Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq18"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq18"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq19 đến Seq25 tương tự như trên...
                        // Và tiếp tục cho tất cả các cột Seq tương ứng...
                        // Xử lý cho cột Seq19
                        object cellValueSeq19 = worksheet.Cells[row, 28].Value;
                        if (cellValueSeq19 != null)
                        {
                            if (decimal.TryParse(cellValueSeq19.ToString(), out decimal seq19Value))
                            {
                                newRow["Seq19"] = seq19Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq19"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq19"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq20 đến Seq25 tương tự như trên...
                        // Size 20
                        object cellValueSeq20 = worksheet.Cells[row, 29].Value;
                        if (cellValueSeq20 != null)
                        {
                            if (decimal.TryParse(cellValueSeq20.ToString(), out decimal seq20Value))
                            {
                                newRow["Seq20"] = seq20Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq20"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq20"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq21 đến Seq25 tương tự như trên...
                        // Và tiếp tục cho tất cả các cột Seq tương ứng...
                        // Xử lý cho cột Seq21
                        object cellValueSeq21 = worksheet.Cells[row, 30].Value;
                        if (cellValueSeq21 != null)
                        {
                            if (decimal.TryParse(cellValueSeq21.ToString(), out decimal seq21Value))
                            {
                                newRow["Seq21"] = seq21Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq21"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq21"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq22 đến Seq25 tương tự như trên...
                        // Size 22
                        object cellValueSeq22 = worksheet.Cells[row, 31].Value;
                        if (cellValueSeq22 != null)
                        {
                            if (decimal.TryParse(cellValueSeq22.ToString(), out decimal seq22Value))
                            {
                                newRow["Seq22"] = seq22Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq22"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq22"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho các cột Seq23 đến Seq25 tương tự như trên...
                        // Và tiếp tục cho tất cả các cột Seq tương ứng...
                        // Xử lý cho cột Seq23
                        object cellValueSeq23 = worksheet.Cells[row, 32].Value;
                        if (cellValueSeq23 != null)
                        {
                            if (decimal.TryParse(cellValueSeq23.ToString(), out decimal seq23Value))
                            {
                                newRow["Seq23"] = seq23Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq23"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq23"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho cột Seq24 tương tự như trên...
                        // Size 24
                        object cellValueSeq24 = worksheet.Cells[row, 33].Value;
                        if (cellValueSeq24 != null)
                        {
                            if (decimal.TryParse(cellValueSeq24.ToString(), out decimal seq24Value))
                            {
                                newRow["Seq24"] = seq24Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq24"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq24"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }

                        // Tiếp tục xử lý cho cột Seq25 tương tự như trên...
                        // Size 25
                        object cellValueSeq25 = worksheet.Cells[row, 34].Value;
                        if (cellValueSeq25 != null)
                        {
                            if (decimal.TryParse(cellValueSeq25.ToString(), out decimal seq25Value))
                            {
                                newRow["Seq25"] = seq25Value;
                            }
                            else
                            {
                                // Xử lý trường hợp khi giá trị không thể chuyển đổi thành Decimal
                                newRow["Seq25"] = 0; // hoặc bất kỳ giá trị mặc định nào phù hợp
                            }
                        }
                        else
                        {
                            newRow["Seq25"] = DBNull.Value; // hoặc null nếu cột cho phép giá trị null
                        }
                        //
                        //
                        //So Chi Lenh
                        string pattenscl = ".*CHI LENH.*";
                        object cellValuescl = worksheet.Cells[row, 35].Value;
                        if (cellValuescl != null && !Regex.IsMatch(cellValuescl.ToString(), pattenscl))
                        {
                            newRow["SO_CHI_LENH"] = cellValuescl.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["SO_CHI_LENH"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //THUC_TE_PC
                        string pattenttpc = ".*THUC TE PC.*";
                        object cellValuettpc = worksheet.Cells[row, 36].Value;
                        if (cellValuettpc != null && !Regex.IsMatch(cellValuettpc.ToString(), pattenttpc))
                        {
                            newRow["THUC_TE_PC"] = cellValuettpc.ToString();
                        }

                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["THUC_TE_PC"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //LUY_TICH_PC
                        string pattenltpc = ".*LUY TICH PC.*";
                        object cellValueltpc = worksheet.Cells[row, 37].Value;
                        
                        if (cellValueltpc != null && !Regex.IsMatch(cellValueltpc.ToString(), pattenltpc))
                        {
                            newRow["LUY_TICH_PC"] = cellValueltpc.ToString();
                            if (cellValueltpc.ToString().Equals("#REF!"))
                            {
                                newRow["LUY_TICH_PC"] = "0";
                            }
                        }
                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["LUY_TICH_PC"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //SO_CHUA_PC
                        string pattenscpc = ".*SO CHUA PC.*";
                        object cellValuelscpc = worksheet.Cells[row, 38].Value;

                        if (cellValuelscpc != null && !Regex.IsMatch(cellValuelscpc.ToString(), pattenscpc))
                        {
                            newRow["SO_CHUA_PC"] = cellValuelscpc.ToString();
                            if (cellValuelscpc.ToString().Equals("#REF!"))
                            {
                                newRow["SO_CHUA_PC"] = "0";
                            }
                        }
                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["SO_CHUA_PC"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }


                        //Don_Vi_San_Xuat
                        string pattendvxs = ".*针车回转表.*";
                        object cellValueldvsx = worksheet.Cells[row, 1].Value;

                        if (cellValueldvsx != null && Regex.IsMatch(cellValueldvsx.ToString(), pattendvxs)) // && !Regex.IsMatch(cellValueldvsx.ToString(), pattendvxs)
                        {

                            //newRow["Don_Vi_San_Xuat"] = cellValueldvsx.ToString();

                            string donViSanXuat = cellValueldvsx.ToString();

                            // Biểu thức chính quy để tìm chuỗi "B1-L" hoặc "B2-L" theo sau bởi các chữ số
                            string pattern = @"B\d+-L\d+";

                            // Sử dụng Regex để tìm chuỗi "B1-L15" trong biến donViSanXuat
                            Match match = Regex.Match(donViSanXuat, pattern);

                            if (match.Success)
                            {
                                // Lấy giá trị từ kết quả tìm kiếm
                                string desiredValue = match.Value;
                                newRow["Don_Vi_San_Xuat"] = match.Value;
                            }
                        }
                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Don_Vi_San_Xuat"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                        //Created_Date
                        string pattendate = ".*制表日期.*";
                        object cellValuedate = worksheet.Cells[row, 1].Value;

                        if (cellValuedate != null && Regex.IsMatch(cellValuedate.ToString(), pattendate)) // && Regex.IsMatch(cellValuedate.ToString(), pattendate)
                        {
                            //newRow["Created_Date"] = cellValuedate.ToString();
                            string dateString = cellValuedate.ToString();
                            // Tách chuỗi bằng dấu cách và lấy phần tử cuối cùng, tức là ngày tháng năm
                            string[] parts = dateString.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            string datePart = parts[parts.Length - 1];

                            // Lấy phần ngày tháng năm từ chuỗi datePart
                            if (DateTime.TryParseExact(datePart, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                            {
                                // Gán giá trị đã chuyển đổi thành ngày tháng năm cho cột "Created_Date"
                                newRow["Created_Date"] = date.ToString("yyyy/MM/dd");
                            }
                        }
                        else
                        {
                            // Xử lý trường hợp khi giá trị của ô là null
                            newRow["Created_Date"] = null; // hoặc bất kỳ giá trị mặc định nào phù hợp
                        }
                    }
                }

                // Thêm hàng mới vào DataTable
                SaveDataToDatabase(excelData);


                // Hiển thị số lượng sheet đã tạo mới trên Label
                //UpdateSheetCountLabel(numberOfSheetsCreated);


                // Hiển thị thông báo thành công
                //MessageBox.Show("Completed");
                //
            }
        }
        
        private void SaveDataToDatabase(DataTable data)
        {
            // Chuẩn bị truy vấn SQL để chèn dữ liệu vào cơ sở dữ liệu
            string query = "INSERT INTO BANG_XOAY_TUA (ProNo, Ten_Giay, Dao_Chat, Article, Dang_Fom, Goo, May, Chat, Ry, size1,size2,size3,size4,size5,size6,size7,size8,size9,size10,size11,size12,size13,size14,size15,size16,size17,size18,size19,size20,size21,size22,size23,size24,size25,SO_CHI_LENH,THUC_TE_PC,LUY_TICH_PC,SO_CHUA_PC,Don_Vi_San_Xuat,Created_Date) VALUES (@Value1, @Value3, @Value4, @Value5, @Value6, @Value7, @Value8, @Value9, @Value10, @Value11,@Value12,@Value13,@Value14,@Value15,@Value16,@Value17,@Value18,@Value19,@Value20,@Value21,@Value22,@Value23,@Value24,@Value25,@Value26,@Value27,@Value28,@Value29,@Value30,@Value31,@Value32,@Value33,@Value34,@Value35,@Value36,@Value37,@Value38,@Value39,@Value40,@Value41)";

            // Lặp qua từng hàng trong DataTable và chèn dữ liệu vào cơ sở dữ liệu
            foreach (DataRow row in data.Rows)
            {
                // Tạo mảng tham số để truyền giá trị vào truy vấn SQL Article
                SqlParameter[] parameters =
                {
            new SqlParameter("@Value1", SqlDbType.VarChar) { Value = row["ProNo"] },
            //new SqlParameter("@Value2", SqlDbType.VarChar) { Value = row["Lean"] },
            new SqlParameter("@Value3", SqlDbType.VarChar) { Value = row["Ten_Giay"] },
            new SqlParameter("@Value4", SqlDbType.VarChar) { Value = row["Dao_Chat"] },
            new SqlParameter("@Value5", SqlDbType.VarChar) { Value = row["Article"] },
            new SqlParameter("@Value6", SqlDbType.VarChar) { Value = row["Dang_Fom"] },
            new SqlParameter("@Value7", SqlDbType.VarChar) { Value = row["Goo"] },
            new SqlParameter("@Value8", SqlDbType.VarChar) { Value = row["May"] },
            new SqlParameter("@Value9", SqlDbType.VarChar) { Value = row["Chat"] },
            new SqlParameter("@Value10", SqlDbType.VarChar) { Value = row["Ry"] },
            new SqlParameter("@Value11", SqlDbType.Decimal) { Value = row["Seq1"] },
            new SqlParameter("@Value12", SqlDbType.Decimal) { Value = row["Seq2"] },
            new SqlParameter("@Value13", SqlDbType.Decimal) { Value = row["Seq3"] },
            new SqlParameter("@Value14", SqlDbType.Decimal) { Value = row["Seq4"] },
            new SqlParameter("@Value15", SqlDbType.Decimal) { Value = row["Seq5"] },
            new SqlParameter("@Value16", SqlDbType.Decimal) { Value = row["Seq6"] },
            new SqlParameter("@Value17", SqlDbType.Decimal) { Value = row["Seq7"] },
            new SqlParameter("@Value18", SqlDbType.Decimal) { Value = row["Seq8"] },
            new SqlParameter("@Value19", SqlDbType.Decimal) { Value = row["Seq9"] },
            new SqlParameter("@Value20", SqlDbType.Decimal) { Value = row["Seq10"] },
            new SqlParameter("@Value21", SqlDbType.Decimal) { Value = row["Seq11"] },
            new SqlParameter("@Value22", SqlDbType.Decimal) { Value = row["Seq12"] },
            new SqlParameter("@Value23", SqlDbType.Decimal) { Value = row["Seq13"] },
            new SqlParameter("@Value24", SqlDbType.Decimal) { Value = row["Seq14"] },
            new SqlParameter("@Value25", SqlDbType.Decimal) { Value = row["Seq15"] },
            new SqlParameter("@Value26", SqlDbType.Decimal) { Value = row["Seq16"] },
            new SqlParameter("@Value27", SqlDbType.Decimal) { Value = row["Seq17"] },
            new SqlParameter("@Value28", SqlDbType.Decimal) { Value = row["Seq18"] },
            new SqlParameter("@Value29", SqlDbType.Decimal) { Value = row["Seq19"] },
            new SqlParameter("@Value30", SqlDbType.Decimal) { Value = row["Seq20"] },
            new SqlParameter("@Value31", SqlDbType.Decimal) { Value = row["Seq21"] },
            new SqlParameter("@Value32", SqlDbType.Decimal) { Value = row["Seq22"] },
            new SqlParameter("@Value33", SqlDbType.Decimal) { Value = row["Seq23"] },
            new SqlParameter("@Value34", SqlDbType.Decimal) { Value = row["Seq24"] },
            new SqlParameter("@Value35", SqlDbType.Decimal) { Value = row["Seq25"] },
            new SqlParameter("@Value36", SqlDbType.VarChar) { Value = row["SO_CHI_LENH"] },
            new SqlParameter("@Value37", SqlDbType.VarChar) { Value = row["THUC_TE_PC"] },
            new SqlParameter("@Value38", SqlDbType.VarChar) { Value = row["LUY_TICH_PC"] },
            new SqlParameter("@Value39", SqlDbType.VarChar) { Value = row["SO_CHUA_PC"] },
            new SqlParameter("@Value40", SqlDbType.VarChar) { Value = row["Don_Vi_San_Xuat"] },
            new SqlParameter("@Value41", SqlDbType.DateTime) { Value = row["Created_Date"] }


            
        };

                // Thêm các tham số cho các cột khác nếu cần thiết 

                // Thực thi truy vấn chèn dữ liệu vào cơ sở dữ liệu
                dbConnect.ExecuteNonQuery(query, parameters);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            // Mặc định đường dẫn vào thư mục C:\ và thư mục ERP và tên tệp Excel "filexoaytua_input.xlsx"
            string defaultPath = @"C:\ERPP\filexoaytua_input.xlsx";

            // Kiểm tra sự tồn tại của tệp Excel
            if (File.Exists(defaultPath))
            {
                ProcessExcelFile(defaultPath);
                // Hiển thị MessageBox
                MessageBox.Show("Finish", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

                // Tạo và cấu hình Timer
                timer = new Timer();
                timer.Interval = 100; // Đặt thời gian đếm là 4 giây (4000 milliseconds)
                timer.Tick += Timer_Tick; // Đăng ký sự kiện Tick
                timer.Start(); // Bắt đầu đếm thời gian
            }
            else
            {
                // Nếu không tìm thấy tệp, hiển thị thông báo
                //MessageBox.Show("", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show("Không tìm thấy tệp Excel filexoaytua_input.xlsx trong thư mục C:\\ERP\\filexoaytua_input", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                timer = new Timer();
                timer.Interval = 100; // Đặt thời gian đếm là 4 giây (4000 milliseconds)
                timer.Tick += Timer_Tick; // Đăng ký sự kiện Tick
                timer.Start(); // Bắt đầu đếm thời gian
            }
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            // Đóng MessageBox khi Timer đếm đủ thời gian
            timer.Stop(); // Dừng Timer
            timer.Dispose(); // Giải phóng Timer
            Application.Exit(); // Đóng ứng dụng
        }
        //


        //
    }
}

