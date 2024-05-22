using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Security.Principal;
using System.Globalization;
using System.Threading;

namespace WindowsFormsApp
{
    public partial class Form1 : Form
    {
        private DBconnect dbConnect;
        private long ProNo; // Bắt đầu từ 0
        string username ="Error";
        public Form1()
        {
            InitializeComponent();
            dbConnect = new DBconnect(); // Khởi tạo kết nối cơ sở dữ liệu
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "IMPORT";

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
                excelData.Columns.Add("Row");
                excelData.Columns.Add("UserID");

                //
                //
                // Lấy giá trị năm và tháng từ thời gian hiện tại

                // Lặp qua từng sheet trong workbook 
                foreach (ExcelWorksheet worksheet in workbook.Worksheets)
                {
                    int currentRow = 0; // Biến đếm hàng hiện tại
                    // Lặp qua từng hàng trong sheet
                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        // Tạo một hàng mới trong DataTable excelData
                        DataRow newRow = excelData.NewRow();


                        //xu ly UserID
                        newRow["UserID"] = username;

                        // Gán giá trị ProNo cho newRow["ProNo"]
                        newRow["ProNo"] = ProNo;

                        //xu ly pro no
                        object cellValue1 = worksheet.Cells[row, 2].Value;
                       //if (cellValue1 != null && !cellValue1.ToString().Equals("型體名稱\nTÊN GIÀY") && !cellValue1.ToString().Equals("合計：") && !cellValue1.ToString().Equals("型\n預\n計\n生\n產\n時\n間"))
                        if (cellValue1 != null && !cellValue1.ToString().Equals("型體名稱\nTÊN GIÀY") && !cellValue1.ToString().Equals("型\n預\n計\n生\n產\n時\n間"))
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



                        // Gán giá trị hàng cho cột "Row"
                        newRow["Row"] = currentRow;

                        // Cập nhật giá trị hàng cho lần lặp tiếp theo
                        currentRow++;

                        //currentRow = 0;  BANG XOAY TUA  ( MAY )
                        object cellValuerow = worksheet.Cells[row, 1].Value;
                        if (cellValuerow != null && cellValuerow.ToString().Equals("BANG XOAY TUA  ( MAY )"))
                        {
                            currentRow = 0;
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
                            newRow["Don_Vi_San_Xuat"] = donViSanXuat;

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
                            else
                            {
                                TextBox textBox = new TextBox
                                {
                                    Text = donViSanXuat,
                                    Multiline = true,
                                    ReadOnly = true,
                                    Dock = DockStyle.Fill,
                                    ScrollBars = ScrollBars.Both,
                                    WordWrap = true
                                };

                                // Tạo một Form để chứa TextBox
                                Form form = new Form
                                {
                                    Text = "Đơn vị sản xuất bị bỏ trống B?-L?: ",
                                    Width = 400,
                                    Height = 200,
                                    StartPosition = FormStartPosition.CenterScreen
                                };

                                // Thêm TextBox vào Form
                                form.Controls.Add(textBox);

                                // Đăng ký sự kiện Click để chọn tất cả văn bản trong TextBox
                                //textBox.Click += (sender, e) => textBox.SelectAll();

                                // Hiển thị Form
                                form.ShowDialog();



                                dbConnect.ExecuteQuery(@"delete BANG_XOAY_TUA");
                                Application.Exit();
                                return;
                            }
                            //string desiredValue = match.Value;
                            //newRow["Don_Vi_San_Xuat"] = match.Value;
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
            string query = "INSERT INTO BANG_XOAY_TUA (ProNo, Ten_Giay, Dao_Chat, Article, Dang_Fom, Goo, May, Chat, Ry, size1,size2,size3,size4,size5,size6,size7,size8,size9,size10,size11,size12,size13,size14,size15,size16,size17,size18,size19,size20,size21,size22,size23,size24,size25,SO_CHI_LENH,THUC_TE_PC,LUY_TICH_PC,SO_CHUA_PC,Don_Vi_San_Xuat,Created_Date, Row, UserID) VALUES (@Value1, @Value3, @Value4, @Value5, @Value6, @Value7, @Value8, @Value9, @Value10, @Value11,@Value12,@Value13,@Value14,@Value15,@Value16,@Value17,@Value18,@Value19,@Value20,@Value21,@Value22,@Value23,@Value24,@Value25,@Value26,@Value27,@Value28,@Value29,@Value30,@Value31,@Value32,@Value33,@Value34,@Value35,@Value36,@Value37,@Value38,@Value39,@Value40,@Value41,@Value42,@Value43)";

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
            new SqlParameter("@Value41", SqlDbType.DateTime) { Value = row["Created_Date"] },
            new SqlParameter("@Value42", SqlDbType.Int) { Value = row["Row"] },
            new SqlParameter("@Value43", SqlDbType.VarChar) { Value = row["UserID"] }


        };

                // Thêm các tham số cho các cột khác nếu cần thiết 

                // Thực thi truy vấn chèn dữ liệu vào cơ sở dữ liệu
                dbConnect.ExecuteNonQuery(query, parameters);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            using (LoginForm loginForm = new LoginForm())
            {
                if (loginForm.ShowDialog() == DialogResult.OK)
                {
                    // Lấy giá trị username từ LoginForm
                    username = loginForm.Username;
                    // Tiếp tục thực hiện công việc khi đăng nhập thành công
                    // Đoạn code ở đây
                    try
                    {
                        // Mặc định đường dẫn vào thư mục C:\ và thư mục ERP và tên tệp Excel "filexoaytua_input.xlsx"
                        //string defaultPath = @"C:\ERPP\filexoaytua_input_chuan1.xlsx";
                        //string defaultPath = @"C:\ERPP\filexoaytua_input.xlsx";


                        string defaultPath = @"C:\ERP\filexoaytua_input.xlsx";
                        // Kiểm tra sự tồn tại của tệp Excel
                        if (File.Exists(defaultPath))
                        {
                            ProcessExcelFile(defaultPath);
                            // Hiển thị MessageBox
                            //MessageBox.Show("Finish", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);


                            //check size
                            CheckSize();
                            //check giờ tua
                            Check_Gio_Tua();
                                //chay procedure

                                // Tạo một đối tượng DBconnect
                                DBconnect dbConnect = new DBconnect();

                                // Tạo một đối tượng SqlCommand
                                using (SqlCommand cmd = new SqlCommand("exec usp_plan_n223_Insert_Xoay_Tua_To_Progress"))
                                {
                                    // Đặt thời gian chờ của truy vấn
                                    cmd.CommandTimeout = 6000; // 10 phút là 600 giây

                                    // Thực thi câu truy vấn bằng phương thức mới
                                    int rowsAffected = dbConnect.ExecuteSqlCommand(cmd);

                                    // Hiển thị thông báo khi thực thi thành công
                                    if (rowsAffected > 0)
                                    {
                                        MessageBox.Show("Finish", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                                        // Đóng ứng dụng
                                        Application.Exit();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Có lỗi xảy ra khi thực thi câu truy vấn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                                        Application.Exit();
                                    }
                                }
 


                        }
                        else
                        {
                            // Nếu không tìm thấy tệp, hiển thị thông báo
                            //MessageBox.Show("", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            MessageBox.Show("Không tìm thấy tệp Excel filexoaytua_input.xlsx trong thư mục C:\\ERP\\filexoaytua_input", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

                            Application.Exit();
                        }
                    }
                    catch (Exception ex)
                    {
                        // Xử lý lỗi và hiển thị thông báo
                        MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                        Application.Exit();
                    }
                }
                else
                {
                    // Người dùng đã hủy đăng nhập hoặc đăng nhập không thành công, 
                    Application.Exit();
                }
            }
            
        }

        //
        /*
        public static void CheckSize()
        {
            DBconnect dbConnect = new DBconnect();
            string sql = @"create table #temp_checksize1(
	id int,
	row int,
	prono varchar(20),
	donvisanxuat varchar(255),
	ry varchar(255),
	tengiay varchar(255),
	dangfom varchar(255),
	qty int,
	size1 DECIMAL(15,1),
    size2 DECIMAL(15,1),
    size3 DECIMAL(15,1),
    size4 DECIMAL(15,1),
    size5 DECIMAL(15,1),
    size6 DECIMAL(15,1),
    size7 DECIMAL(15,1),
    size8 DECIMAL(15,1),
    size9 DECIMAL(15,1),
    size10 DECIMAL(15,1),
    size11 DECIMAL(15,1),
    size12 DECIMAL(15,1),
    size13 DECIMAL(15,1),
    size14 DECIMAL(15,1),
    size15 DECIMAL(15,1),
    size16 DECIMAL(15,1),
    size17 DECIMAL(15,1),
    size18 DECIMAL(15,1),
    size19 DECIMAL(15,1),
    size20 DECIMAL(15,1),
    size21 DECIMAL(15,1),
    size22 DECIMAL(15,1),
    size23 DECIMAL(15,1),
    size24 DECIMAL(15,1),
    size25 DECIMAL(15,1),
	UserID varchar(20)
	)
;
--
-- them du lieu vao bang #temp_checksize1
INSERT INTO #temp_checksize1 (ProNo,row,donvisanxuat,ry,tengiay,dangfom, size1,size2,size3,size4,size5,size6,size7,size8,size9,size10,size11,size12,size13,size14,size15,size16,size17,size18,size19,size20,size21,size22,size23,size24,size25,UserID)
select ProNo,row,Don_Vi_San_Xuat,ry,Ten_Giay,Dang_Fom,size1,size2,size3,size4,size5,size6,size7,size8,size9,size10,size11,size12,size13,size14,size15,size16,size17,size18,size19,size20,size21,size22,size23,size24,size25,UserID
from bang_xoay_tua
ORDER BY ProNo, row ASC
;
--update donvisanxuat bang #temp_checksize1
update #temp_checksize1
set donvisanxuat = bang_xoay_tua.Don_Vi_San_Xuat
from #temp_checksize1
join bang_xoay_tua on bang_xoay_tua.prono=#temp_checksize1.prono
;
-- update dangfom tu lastname cua bang lastnom cho bang #temp_checksize1
update #temp_checksize1
set dangfom = LastNoM.LastName
from #temp_checksize1
join bang_xoay_tua on bang_xoay_tua.prono=#temp_checksize1.prono
join LastNoM on LastNoM.LastNo=bang_xoay_tua.Dang_Fom
;
--delete cac 2 hang dau de cat size
delete #temp_checksize1
where row in (0,1)
;
Create Table #temp_checksize2(
	id int,
	row int,
	prono varchar(20),
	dvsx varchar(255),
	ry varchar(255),
	dangfom varchar(255),
	size DECIMAL(15,1),
	qty int,
	UserID varchar(20)
)
;
--them size theo thu tu 1-25 cot size = 1 ma prono
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, donvisanxuat, ry, dangfom,UserID, Size
FROM (
    SELECT row, prono, donvisanxuat, ry, dangfom,UserID,Size,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM #temp_checksize1
    UNPIVOT (
        Size FOR SizeNumber IN (
            size1, size2, size3, size4, size5, 
            size6, size7, size8, size9, size10,
            size11, size12, size13, size14, size15,
            size16, size17, size18, size19, size20,
            size21, size22, size23, size24, size25
        )
    ) AS unpvt
) AS NumberedRows
WHERE RowNum <= 25
ORDER BY prono, donvisanxuat;
;
--lap lai bang #temp_checksize2 de mo rong prono: VD prono1 co 4 lenh tuong duong 25x4=100 hang prono1
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
------------------------------------------------------------
CREATE TABLE #temp_checksize3 (
    ry VARCHAR(255),
    prono VARCHAR(255),
	id int,
)
;
--them nhung lenh co so hang nho hon so hang gap dieu kien '合計：'-->trong file excel thi se add vo bang #tempdata
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY #temp_checksize1.prono ORDER BY #temp_checksize1.[row]) AS RowNumber
    FROM #temp_checksize1
)
INSERT INTO #temp_checksize3 (ry, prono)
SELECT NumberedRows.ry, NumberedRows.prono
FROM NumberedRows
INNER JOIN (
    SELECT MAX(#temp_checksize1.[row]) AS max_row, #temp_checksize1.prono
    FROM #temp_checksize1
    WHERE tengiay = '合計：'
    GROUP BY #temp_checksize1.prono
) AS MaxRows ON NumberedRows.prono = MaxRows.prono
WHERE NumberedRows.RowNumber <= MaxRows.max_row;

;
--update id 1 - ... #temp_checksize2  1 id tuong duong 25 size = 25 hang
WITH CTE AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS row_num,
           (ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) - 1) / 25 + 1 AS group_num
    FROM #temp_checksize2
)
UPDATE CTE
SET id = group_num
;
delete #temp_checksize3 where ry is null
;
--update id 1 - ... #TempData danh so id cho bang #tempdata
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM #temp_checksize3
)
UPDATE NumberedRows
SET ID = RowNum
;
-- so sanh neu prono va id bang nhau them them vao (id trong bang #tempdata la nhung hang nam truoc ky tu '合計：' trong excel) --> thoa man dieu kien
update #temp_checksize2
 set ry = #temp_checksize3.ry
 from #temp_checksize3
 where #temp_checksize3.id=#temp_checksize2.id and #temp_checksize3.prono=#temp_checksize2.prono
;
--xoa ry null nhung hang bi du thua.
delete #temp_checksize2 where ry is null
;
-- Cat hang cho bang voi hang trong bang #tempdata
delete #temp_checksize1 where ry is null
;
--update id trong bang #temp_checksize1
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM #temp_checksize1
)
UPDATE NumberedRows
SET ID = RowNum
;
--update size
UPDATE #temp_checksize2
SET qty = 
    CASE 
        WHEN #temp_checksize2.size = 3.0 THEN #temp_checksize1.size1
        WHEN #temp_checksize2.size = 3.5 THEN #temp_checksize1.size2
        WHEN #temp_checksize2.size = 4.0 THEN #temp_checksize1.size3
		WHEN #temp_checksize2.size = 4.5 THEN #temp_checksize1.size4
		WHEN #temp_checksize2.size = 5.0 THEN #temp_checksize1.size5
		WHEN #temp_checksize2.size = 5.5 THEN #temp_checksize1.size6
		WHEN #temp_checksize2.size = 6.0 THEN #temp_checksize1.size7
		WHEN #temp_checksize2.size = 6.5 THEN #temp_checksize1.size8
		WHEN #temp_checksize2.size = 7.0 THEN #temp_checksize1.size9
		WHEN #temp_checksize2.size = 7.5 THEN #temp_checksize1.size10
		WHEN #temp_checksize2.size = 8.0 THEN #temp_checksize1.size11
		WHEN #temp_checksize2.size = 8.5 THEN #temp_checksize1.size12
		WHEN #temp_checksize2.size = 9.0 THEN #temp_checksize1.size13
		WHEN #temp_checksize2.size = 9.5 THEN #temp_checksize1.size14
		WHEN #temp_checksize2.size = 10.0 THEN #temp_checksize1.size15
		WHEN #temp_checksize2.size = 10.5 THEN #temp_checksize1.size16
		WHEN #temp_checksize2.size = 11.0 THEN #temp_checksize1.size17
		WHEN #temp_checksize2.size = 11.5 THEN #temp_checksize1.size18
		WHEN #temp_checksize2.size = 12.0 THEN #temp_checksize1.size19
		WHEN #temp_checksize2.size = 12.5 THEN #temp_checksize1.size20
		WHEN #temp_checksize2.size = 13.0 THEN #temp_checksize1.size21
		WHEN #temp_checksize2.size = 13.5 THEN #temp_checksize1.size22
		WHEN #temp_checksize2.size = 14.0 THEN #temp_checksize1.size23
		WHEN #temp_checksize2.size = 14.5 THEN #temp_checksize1.size24
		WHEN #temp_checksize2.size = 15.0 THEN #temp_checksize1.size25
    END
FROM #temp_checksize2
JOIN #temp_checksize1 ON #temp_checksize1.prono = #temp_checksize2.prono AND #temp_checksize1.id = #temp_checksize2.id;
;
--xoa nhung size co qty la null
Delete from #temp_checksize2
where qty is null
;
--xoa nhung hang co ca 3 cot bi trung lap va giu lai 1 cot
WITH cte AS (
  SELECT dangfom, ry, size, ROW_NUMBER() OVER (PARTITION BY dangfom, ry, size ORDER BY (SELECT NULL)) AS rn
  FROM #temp_checksize2
)
DELETE FROM cte
WHERE rn > 1;
;
ALTER TABLE #temp_checksize2
ALTER COLUMN size VARCHAR(20)
;
UPDATE #temp_checksize2
SET size = 
    CASE 
        WHEN CHARINDEX('.', CAST(size AS VARCHAR(20))) = 2 AND RIGHT(CAST(size AS VARCHAR(20)), 2) = '.0' THEN '0' + LEFT(CAST(size AS VARCHAR(20)), 1)
        WHEN CHARINDEX('.', CAST(size AS VARCHAR(20))) = 2 THEN '0' + CAST(size AS VARCHAR(20))
        WHEN CHARINDEX('.', CAST(size AS VARCHAR(20))) = 3 AND RIGHT(CAST(size AS VARCHAR(20)), 2) = '.0' THEN LEFT(CAST(size AS VARCHAR(20)), 2)
        ELSE CAST(size AS VARCHAR(20))
    END
WHERE size IS NOT NULL
;
delete #temp_checksize2 where qty=0
;
-- Tạo một bảng tạm để lưu kết quả so sánh
CREATE TABLE #MismatchRy(
    ry VARCHAR(255),
    CountInTemptableas2 INT,
    CountInDdzls INT
);

-- Chèn vào bảng tạm các giá trị ry có số lượng khác nhau giữa hai bảng
INSERT INTO #MismatchRy(ry, CountInTemptableas2, CountInDdzls)
SELECT t1.ry, t1.CountInTemptableas2, t2.CountInDdzls
FROM (
    SELECT ry, COUNT(*) AS CountInTemptableas2
    FROM #temp_checksize2
    GROUP BY ry
) AS t1
LEFT JOIN (
    SELECT DDBH AS ry, COUNT(*) AS CountInDdzls
    FROM ddzls
    GROUP BY DDBH
) AS t2
ON t1.ry = t2.ry
WHERE t1.CountInTemptableas2 <> t2.CountInDdzls
;
SELECT COUNT(*) AS InvalidSizeCount
    FROM #MismatchRy
;
Drop table #temp_checksize1
;
Drop table #temp_checksize2
;
drop table #temp_checksize3
;
drop table #MismatchRy
;
delete BANG_XOAY_TUA";

            try
            {
                int invalidSizeCount = dbConnect.ExecuteScalarInt(sql);

                if (invalidSizeCount > 0)
                {
                    MessageBox.Show("Thông báo lỗi: Có giá trị size không hợp lệ.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (FormatException ex)
            {
                MessageBox.Show("Đã xảy ra lỗi định dạng khi thực hiện kiểm tra kích thước size: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi khi thực hiện kiểm tra kích thước size: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        */
        //
        public static void CheckSize()
        {
            DBconnect dbConnect = new DBconnect();
            string sql = @"
update a
set a.Don_Vi_San_Xuat = b.Don_Vi_San_Xuat
from BANG_XOAY_TUA a
join BANG_XOAY_TUA b on a.ProNo=b.ProNo
--where a.Don_Vi_San_Xuat is null and b.Don_Vi_San_Xuat is not null
where (a.Don_Vi_San_Xuat is null) or (LEN(LTRIM(RTRIM(a.Don_Vi_San_Xuat))) = 0) and (LEN(LTRIM(RTRIM(b.Don_Vi_San_Xuat))) > 0)
IF OBJECT_ID('#temp_checksize1', 'U') IS NOT NULL
BEGIN
    DROP TABLE #temp_checksize1;
END

IF OBJECT_ID('#temp_checksize2', 'U') IS NOT NULL
BEGIN
    DROP TABLE #temp_checksize2;
END

IF OBJECT_ID('#temp_checksize3', 'U') IS NOT NULL
BEGIN
    DROP TABLE #temp_checksize3;
END

IF OBJECT_ID('#temp_checksize4', 'U') IS NOT NULL
BEGIN
    DROP TABLE #temp_checksize4;
END
create table #temp_checksize1(
	id int,
	row int,
	prono varchar(20),
	donvisanxuat varchar(255),
	ry varchar(255),
	tengiay varchar(255),
	dangfom varchar(255),
	qty int,
	size1 DECIMAL(15,1),
    size2 DECIMAL(15,1),
    size3 DECIMAL(15,1),
    size4 DECIMAL(15,1),
    size5 DECIMAL(15,1),
    size6 DECIMAL(15,1),
    size7 DECIMAL(15,1),
    size8 DECIMAL(15,1),
    size9 DECIMAL(15,1),
    size10 DECIMAL(15,1),
    size11 DECIMAL(15,1),
    size12 DECIMAL(15,1),
    size13 DECIMAL(15,1),
    size14 DECIMAL(15,1),
    size15 DECIMAL(15,1),
    size16 DECIMAL(15,1),
    size17 DECIMAL(15,1),
    size18 DECIMAL(15,1),
    size19 DECIMAL(15,1),
    size20 DECIMAL(15,1),
    size21 DECIMAL(15,1),
    size22 DECIMAL(15,1),
    size23 DECIMAL(15,1),
    size24 DECIMAL(15,1),
    size25 DECIMAL(15,1),
	UserID varchar(20)
	)
;
--
-- them du lieu vao bang #temp_checksize1
INSERT INTO #temp_checksize1 (ProNo,row,donvisanxuat,ry,tengiay,dangfom, size1,size2,size3,size4,size5,size6,size7,size8,size9,size10,size11,size12,size13,size14,size15,size16,size17,size18,size19,size20,size21,size22,size23,size24,size25,UserID)
select ProNo,row,Don_Vi_San_Xuat,ry,Ten_Giay,Dang_Fom,size1,size2,size3,size4,size5,size6,size7,size8,size9,size10,size11,size12,size13,size14,size15,size16,size17,size18,size19,size20,size21,size22,size23,size24,size25,UserID
from bang_xoay_tua
ORDER BY ProNo, row ASC
;
--update donvisanxuat bang #temp_checksize1
update #temp_checksize1
set donvisanxuat = bang_xoay_tua.Don_Vi_San_Xuat
from #temp_checksize1
join bang_xoay_tua on bang_xoay_tua.prono=#temp_checksize1.prono
;
-- update dangfom tu lastname cua bang lastnom cho bang #temp_checksize1
update #temp_checksize1
set dangfom = LastNoM.LastName
from #temp_checksize1
join bang_xoay_tua on bang_xoay_tua.prono=#temp_checksize1.prono
join LastNoM on LastNoM.LastNo=bang_xoay_tua.Dang_Fom
;
--delete cac 2 hang dau de cat size
delete #temp_checksize1
where row in (0,1)
;
Create Table #temp_checksize2(
	id int,
	row int,
	prono varchar(20),
	dvsx varchar(255),
	ry varchar(255),
	dangfom varchar(255),
	size DECIMAL(15,1),
	qty int,
	UserID varchar(20)
)
;
--them size theo thu tu 1-25 cot size = 1 ma prono
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, donvisanxuat, ry, dangfom,UserID, Size
FROM (
    SELECT row, prono, donvisanxuat, ry, dangfom,UserID,Size,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM #temp_checksize1
    UNPIVOT (
        Size FOR SizeNumber IN (
            size1, size2, size3, size4, size5, 
            size6, size7, size8, size9, size10,
            size11, size12, size13, size14, size15,
            size16, size17, size18, size19, size20,
            size21, size22, size23, size24, size25
        )
    ) AS unpvt
) AS NumberedRows
WHERE RowNum <= 25
ORDER BY prono, donvisanxuat;
;
--lap lai bang #temp_checksize2 de mo rong prono: VD prono1 co 4 lenh tuong duong 25x4=100 hang prono1
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
INSERT INTO #temp_checksize2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM #temp_checksize2;
;
------------------------------------------------------------
CREATE TABLE #temp_checksize3 (
    ry VARCHAR(255),
    prono VARCHAR(255),
	id int,
)
;
--them nhung lenh co so hang nho hon so hang gap dieu kien '合計：'-->trong file excel thi se add vo bang #tempdata
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY #temp_checksize1.prono ORDER BY #temp_checksize1.[row]) AS RowNumber
    FROM #temp_checksize1
)
INSERT INTO #temp_checksize3 (ry, prono)
SELECT NumberedRows.ry, NumberedRows.prono
FROM NumberedRows
INNER JOIN (
    SELECT MAX(#temp_checksize1.[row]) AS max_row, #temp_checksize1.prono
    FROM #temp_checksize1
    WHERE tengiay = '合計：'
    GROUP BY #temp_checksize1.prono
) AS MaxRows ON NumberedRows.prono = MaxRows.prono
WHERE NumberedRows.RowNumber <= MaxRows.max_row;

;
--update id 1 - ... #temp_checksize2  1 id tuong duong 25 size = 25 hang
WITH CTE AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS row_num,
           (ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) - 1) / 25 + 1 AS group_num
    FROM #temp_checksize2
)
UPDATE CTE
SET id = group_num
;
delete #temp_checksize3 where ry is null
;
--update id 1 - ... #TempData danh so id cho bang #tempdata
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM #temp_checksize3
)
UPDATE NumberedRows
SET ID = RowNum
;
-- so sanh neu prono va id bang nhau them them vao (id trong bang #tempdata la nhung hang nam truoc ky tu '合計：' trong excel) --> thoa man dieu kien
update #temp_checksize2
 set ry = #temp_checksize3.ry
 from #temp_checksize3
 where #temp_checksize3.id=#temp_checksize2.id and #temp_checksize3.prono=#temp_checksize2.prono
;
--xoa ry null nhung hang bi du thua.
delete #temp_checksize2 where ry is null
;
-- Cat hang cho bang voi hang trong bang #tempdata
delete #temp_checksize1 where ry is null
;
--update id trong bang #temp_checksize1
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM #temp_checksize1
)
UPDATE NumberedRows
SET ID = RowNum
;
--update size
UPDATE #temp_checksize2
SET qty = 
    CASE 
        WHEN #temp_checksize2.size = 3.0 THEN #temp_checksize1.size1
        WHEN #temp_checksize2.size = 3.5 THEN #temp_checksize1.size2
        WHEN #temp_checksize2.size = 4.0 THEN #temp_checksize1.size3
		WHEN #temp_checksize2.size = 4.5 THEN #temp_checksize1.size4
		WHEN #temp_checksize2.size = 5.0 THEN #temp_checksize1.size5
		WHEN #temp_checksize2.size = 5.5 THEN #temp_checksize1.size6
		WHEN #temp_checksize2.size = 6.0 THEN #temp_checksize1.size7
		WHEN #temp_checksize2.size = 6.5 THEN #temp_checksize1.size8
		WHEN #temp_checksize2.size = 7.0 THEN #temp_checksize1.size9
		WHEN #temp_checksize2.size = 7.5 THEN #temp_checksize1.size10
		WHEN #temp_checksize2.size = 8.0 THEN #temp_checksize1.size11
		WHEN #temp_checksize2.size = 8.5 THEN #temp_checksize1.size12
		WHEN #temp_checksize2.size = 9.0 THEN #temp_checksize1.size13
		WHEN #temp_checksize2.size = 9.5 THEN #temp_checksize1.size14
		WHEN #temp_checksize2.size = 10.0 THEN #temp_checksize1.size15
		WHEN #temp_checksize2.size = 10.5 THEN #temp_checksize1.size16
		WHEN #temp_checksize2.size = 11.0 THEN #temp_checksize1.size17
		WHEN #temp_checksize2.size = 11.5 THEN #temp_checksize1.size18
		WHEN #temp_checksize2.size = 12.0 THEN #temp_checksize1.size19
		WHEN #temp_checksize2.size = 12.5 THEN #temp_checksize1.size20
		WHEN #temp_checksize2.size = 13.0 THEN #temp_checksize1.size21
		WHEN #temp_checksize2.size = 13.5 THEN #temp_checksize1.size22
		WHEN #temp_checksize2.size = 14.0 THEN #temp_checksize1.size23
		WHEN #temp_checksize2.size = 14.5 THEN #temp_checksize1.size24
		WHEN #temp_checksize2.size = 15.0 THEN #temp_checksize1.size25
    END
FROM #temp_checksize2
JOIN #temp_checksize1 ON #temp_checksize1.prono = #temp_checksize2.prono AND #temp_checksize1.id = #temp_checksize2.id;
;
--xoa nhung size co qty la null
Delete from #temp_checksize2
where qty is null
;
ALTER TABLE #temp_checksize2
ALTER COLUMN size VARCHAR(20)
;
UPDATE #temp_checksize2
SET size = 
    CASE 
        WHEN CHARINDEX('.', CAST(size AS VARCHAR(20))) = 2 AND RIGHT(CAST(size AS VARCHAR(20)), 2) = '.0' THEN '0' + LEFT(CAST(size AS VARCHAR(20)), 1)
        WHEN CHARINDEX('.', CAST(size AS VARCHAR(20))) = 2 THEN '0' + CAST(size AS VARCHAR(20))
        WHEN CHARINDEX('.', CAST(size AS VARCHAR(20))) = 3 AND RIGHT(CAST(size AS VARCHAR(20)), 2) = '.0' THEN LEFT(CAST(size AS VARCHAR(20)), 2)
        ELSE CAST(size AS VARCHAR(20))
    END
WHERE size IS NOT NULL
;
delete #temp_checksize2 where qty=0
;

-- NẾU TRÙNG RY TRÙNG SIZE, SẼ XÓA RY NÀO SUM QTY LẠI NHỎ HƠN RY CÒN LẠI
-- Bước 1: Tính tổng số lượng (qty) theo prono
WITH TotalQtyByProno AS (
    SELECT prono, dangfom, ry, size,
           SUM(qty) AS total_qty
    FROM #temp_checksize2
    GROUP BY prono, dangfom, ry, size
),
-- Bước 2: Xác định prono có tổng số lượng lớn nhất cho mỗi cặp dangfom, ry, size
MaxTotalQty AS (
    SELECT prono, ry, size,
           MAX(total_qty) AS max_total_qty
    FROM TotalQtyByProno
    GROUP BY prono, ry, size
),
-- Bước 3: Chọn các hàng cần giữ lại (các hàng có prono với tổng số lượng lớn nhất)
RowsToKeep AS (
    SELECT t.prono, t.dangfom, t.ry, t.size, t.qty
    FROM #temp_checksize2 t
    INNER JOIN TotalQtyByProno tq ON t.prono = tq.prono AND t.dangfom = tq.dangfom AND t.ry = tq.ry AND t.size = tq.size
    INNER JOIN MaxTotalQty mt ON tq.prono = mt.prono AND tq.ry = mt.ry AND tq.size = mt.size AND tq.total_qty = mt.max_total_qty
)
-- Bước 4: Xóa các hàng không thuộc RowsToKeep
DELETE t
FROM #temp_checksize2 t
LEFT JOIN RowsToKeep k
ON t.prono = k.prono AND t.dangfom = k.dangfom AND t.ry = k.ry AND t.size = k.size AND t.qty = k.qty
WHERE k.prono IS NULL;

DELETE FROM #temp_checksize2
WHERE EXISTS (
    SELECT 1
    FROM #temp_checksize2 AS t2
    WHERE #temp_checksize2.size = t2.size
    AND #temp_checksize2.ry = t2.ry
    AND #temp_checksize2.Id > t2.Id
);
--xóa trùng size trùng lệnh
WITH CTE AS (
    SELECT 
        size, 
        ry, 
        ROW_NUMBER() OVER (PARTITION BY size, ry ORDER BY (SELECT NULL)) AS row_num
    FROM #temp_checksize2
)
DELETE FROM CTE
WHERE row_num > 1;
-- Xóa -1 -2 ở cuối lệnh
UPDATE #temp_checksize2
SET ry = CASE
            WHEN PATINDEX('%-[0-9]', ry) > 0 AND RIGHT(ry, 2) LIKE '-[0-9]' THEN LEFT(ry, LEN(ry) - 2)
            ELSE ry
         END
WHERE ry LIKE '%-%';
--Tạo 1 bảng để lưu ry: trong file excel có Ry và size đó nhưng trong bảng ddzls không có, và số lượng của size trong excel lớn hơn số lượng của size trong ddzls.
create table #temp_checksize4(
	ry varchar(100),
	size varchar(20)
);
--check size có trong excel nhưng không có trong ddzls, ry có trong excel nhưng không có trong ddzls, số lượng size đó trong excel lớn hơn trong ddzls
insert into #temp_checksize4(ry,size)
SELECT 
    t.ry as ryexcel,
    t.size as sizeexcel
    --t.qty as qtyexcel,
    --ISNULL(d.Quantity, 0) as qtyddzl
FROM 
    #temp_checksize2 t
LEFT JOIN 
    ddzls d ON t.ry = d.DDBH AND t.size = d.cc
WHERE 
    d.cc IS NULL OR t.qty > d.Quantity;
--
SELECT 
    ry, 
    STUFF((
        SELECT ', ' + size
        FROM #temp_checksize4 t2
        WHERE t2.ry = t1.ry
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') as size
FROM 
    #temp_checksize4 t1
GROUP BY 
    ry;

Drop table #temp_checksize1
;
Drop table #temp_checksize2
;
drop table #temp_checksize3
;
drop table #temp_checksize4
;
";

            try
            {
                DataTable resultTable = dbConnect.ExecuteQuery(sql);

                if (resultTable.Rows.Count > 0)
                {
                    //Nếu lỗi thì xóa bang xoay tua để chạy procedure không bị lỗi
                    dbConnect.ExecuteQuery(@"delete BANG_XOAY_TUA");

                    //string message = "Thông báo: Có vấn đề về size ở các lệnh:\r\n";
                    string message = "";

                    // Duyệt qua từng hàng và lấy giá trị từ các cột
                    foreach (DataRow row in resultTable.Rows)
                    {
                        // Giả sử bạn muốn hiển thị giá trị của cột đầu tiên
                        // Bạn có thể thay đổi theo yêu cầu cụ thể của bạn
                        message += row[0].ToString() + ": ";
                        message += row[1].ToString() + "\r\n";
                    }

                    // Tạo và cấu hình một TextBox để hiển thị thông báo
                    TextBox textBox = new TextBox
                    {
                        Text = message,
                        Multiline = true,
                        ReadOnly = true,
                        Dock = DockStyle.Fill,
                        ScrollBars = ScrollBars.Both,
                        WordWrap = true
                    };

                    // Tạo một Form để chứa TextBox
                    Form form = new Form
                    {
                        Text = "Có vấn đề về size ở các lệnh:",
                        Width = 600,
                        Height = 400,
                        StartPosition = FormStartPosition.CenterScreen
                    };

                    // Thêm TextBox vào Form
                    form.Controls.Add(textBox);

                    // Đăng ký sự kiện Click để chọn tất cả văn bản trong TextBox
                    //textBox.Click += (sender, e) => textBox.SelectAll();

                    // Hiển thị Form
                    form.ShowDialog();

                    // Thoát ứng dụng sau khi form được đóng
                    Application.Exit();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi khi thực hiện kiểm tra size: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void Check_Gio_Tua()
        {
            DBconnect dbConnect = new DBconnect();
            string sql = @"
IF OBJECT_ID('temptablell', 'U') IS NOT NULL
BEGIN
    DROP TABLE temptablell;
END
IF OBJECT_ID('temptableas2', 'U') IS NOT NULL
BEGIN
    DROP TABLE temptableas2;
END
IF OBJECT_ID('#TempData', 'U') IS NOT NULL
BEGIN
    DROP TABLE #TempData;
END
IF OBJECT_ID('#TempData2', 'U') IS NOT NULL
BEGIN
    DROP TABLE #TempData2;
END
IF OBJECT_ID('TempTable', 'U') IS NOT NULL
BEGIN
    DROP TABLE TempTable;
END
IF OBJECT_ID('#TempData1', 'U') IS NOT NULL
BEGIN
    DROP TABLE #TempData1;
END
IF OBJECT_ID('#TempData3', 'U') IS NOT NULL
BEGIN
    DROP TABLE #TempData3;
END
IF OBJECT_ID('#TempData4', 'U') IS NOT NULL
BEGIN
    DROP TABLE #TempData4;
END
IF OBJECT_ID('tabletempaskask', 'U') IS NOT NULL
BEGIN
    DROP TABLE tabletempaskask;
END
IF OBJECT_ID('#TempData5', 'U') IS NOT NULL
BEGIN
    DROP TABLE #TempData5;
END
IF OBJECT_ID('#TempData22', 'U') IS NOT NULL
BEGIN
    DROP TABLE #TempData22;
END
create table temptablell(
	id int,
	row int,
	prono varchar(20),
	donvisanxuat varchar(255),
	ry varchar(255),
	tengiay varchar(255),
	dangfom varchar(255),
	qty int,
	sochuapc varchar(50),
	goo varchar(20),
	may varchar(20),
	size1 DECIMAL(15,1),
    size2 DECIMAL(15,1),
    size3 DECIMAL(15,1),
    size4 DECIMAL(15,1),
    size5 DECIMAL(15,1),
    size6 DECIMAL(15,1),
    size7 DECIMAL(15,1),
    size8 DECIMAL(15,1),
    size9 DECIMAL(15,1),
    size10 DECIMAL(15,1),
    size11 DECIMAL(15,1),
    size12 DECIMAL(15,1),
    size13 DECIMAL(15,1),
    size14 DECIMAL(15,1),
    size15 DECIMAL(15,1),
    size16 DECIMAL(15,1),
    size17 DECIMAL(15,1),
    size18 DECIMAL(15,1),
    size19 DECIMAL(15,1),
    size20 DECIMAL(15,1),
    size21 DECIMAL(15,1),
    size22 DECIMAL(15,1),
    size23 DECIMAL(15,1),
    size24 DECIMAL(15,1),
    size25 DECIMAL(15,1),
	UserID varchar(20)
	)
;
-- them du lieu vao bang temptablell
INSERT INTO temptablell (ProNo,row,donvisanxuat,ry,tengiay,dangfom,sochuapc,goo,may, size1,size2,size3,size4,size5,size6,size7,size8,size9,size10,size11,size12,size13,size14,size15,size16,size17,size18,size19,size20,size21,size22,size23,size24,size25,UserID)
select ProNo,row,Don_Vi_San_Xuat,ry,Ten_Giay,Dang_Fom,SO_CHUA_PC,Goo,May,size1,size2,size3,size4,size5,size6,size7,size8,size9,size10,size11,size12,size13,size14,size15,size16,size17,size18,size19,size20,size21,size22,size23,size24,size25,UserID
from bang_xoay_tua
ORDER BY ProNo, row ASC
;
--update donvisanxuat bang temptablell
update temptablell
set donvisanxuat = bang_xoay_tua.Don_Vi_San_Xuat
from temptablell
join bang_xoay_tua on bang_xoay_tua.prono=temptablell.prono
;
--delete cac 2 hang dau de cat size
delete temptablell
where row in (0,1)
;
Create Table temptableas2(
	id int,
	row int,
	prono varchar(20),
	dvsx varchar(255),
	ry varchar(255),
	seq int,
	dangfom varchar(255),
	size DECIMAL(15,1),
	qty int,
	goo varchar(50),
	may varchar(50),
	UserID varchar(20)
)
;
--them size theo thu tu 1-25 cot size = 1 ma prono
INSERT INTO temptableas2 (row, prono, dvsx, ry, dangfom,UserID, size,goo,may)
SELECT row, prono, donvisanxuat, ry, dangfom,UserID, Size,goo,may
FROM (
    SELECT row, prono, donvisanxuat, ry, dangfom,UserID,Size,goo,may,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM temptablell
    UNPIVOT (
        Size FOR SizeNumber IN (
            size1, size2, size3, size4, size5, 
            size6, size7, size8, size9, size10,
            size11, size12, size13, size14, size15,
            size16, size17, size18, size19, size20,
            size21, size22, size23, size24, size25
        )
    ) AS unpvt
) AS NumberedRows
WHERE RowNum <= 25
ORDER BY prono, donvisanxuat;
;
--lap lai bang temptableas2 de mo rong prono: VD prono1 co 4 lenh tuong duong 25x4=100 hang prono1
INSERT INTO temptableas2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM temptableas2;
;
INSERT INTO temptableas2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM temptableas2;
;
INSERT INTO temptableas2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM temptableas2;
;
INSERT INTO temptableas2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM temptableas2;
;
INSERT INTO temptableas2 (row, prono, dvsx, ry, dangfom,UserID, size)
SELECT row, prono, dvsx, ry, dangfom,UserID, Size
FROM temptableas2;
;
------------------------------------------------------------
CREATE TABLE #TempData5 (
    ry VARCHAR(255),
    prono VARCHAR(255),
	id int,
)
;
CREATE TABLE #TempData22 (
    ry VARCHAR(255),
    prono VARCHAR(255),
	id int,
)
;
--them nhung lenh co so hang nho hon so hang gap dieu kien '合計：'-->trong file excel thi se add vo bang #tempdata
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY temptablell.prono ORDER BY temptablell.[row]) AS RowNumber
    FROM temptablell
)
INSERT INTO #TempData5 (ry, prono)
SELECT NumberedRows.ry, NumberedRows.prono
FROM NumberedRows
INNER JOIN (
    SELECT MAX(temptablell.[row]) AS max_row, temptablell.prono
    FROM temptablell
    WHERE tengiay = '合計：'
    GROUP BY temptablell.prono
) AS MaxRows ON NumberedRows.prono = MaxRows.prono
WHERE NumberedRows.RowNumber >= MaxRows.max_row;
;
--lay id de xoa may cai nguoc lai
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY temptablell.prono ORDER BY temptablell.[row]) AS RowNumber
    FROM temptablell
)
INSERT INTO #TempData22 (ry, prono)
SELECT NumberedRows.ry, NumberedRows.prono
FROM NumberedRows
INNER JOIN (
    SELECT MAX(temptablell.[row]) AS max_row, temptablell.prono
    FROM temptablell
    WHERE tengiay = '合計：'
    GROUP BY temptablell.prono
) AS MaxRows ON NumberedRows.prono = MaxRows.prono
WHERE NumberedRows.RowNumber <= MaxRows.max_row;

;
--update id 1 - ... temptableas2  1 id tuong duong 25 size = 25 hang
WITH CTE AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS row_num,
           (ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) - 1) / 25 + 1 AS group_num
    FROM temptableas2
)
UPDATE CTE
SET id = group_num
;
delete #TempData5 where ry is null
;
--
WITH CTE AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS row_num,
           (ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) - 1) / 25 + 1 AS group_num
    FROM temptableas2
)
UPDATE CTE
SET id = group_num
;
delete #TempData22 where ry is null
;
--update id 1 - ... #TempData danh so id cho bang #tempdata
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM #TempData5
)
UPDATE NumberedRows
SET ID = RowNum
;
--update id 1 - ... #TempData danh so id cho bang #tempdata
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM #TempData22
)
UPDATE NumberedRows
SET ID = RowNum
;
-- so sanh neu prono va id bang nhau them them vao (id trong bang #tempdata la nhung hang nam truoc ky tu '合計：' trong excel) --> thoa man dieu kien
update temptableas2
 set ry = #TempData5.ry
 from #TempData5
 where #TempData5.id=temptableas2.id and #TempData5.prono=temptableas2.prono
;
--XOA HANG DAU CHO NGANG VOI BANG #TempData2
DELETE temptablell WHERE ROW  = 2;
--update id trong bang temptablell
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY (SELECT NULL)) AS RowNum
    FROM temptablell
)
UPDATE NumberedRows
SET ID = RowNum
;
--xoa id truoc ky tu ### cat hang                  TOI DAY LA DA CAT HANG KHUC TREN CHECKLLAST ROI, PHAI TINH TONG TRUOC KHI CAT
DELETE temptablell
FROM temptablell
INNER JOIN #TempData22 ON #TempData22.prono = temptablell.prono AND #TempData22.id = temptablell.id;
;
--update Ry cho phu deu 2 hang A B
UPDATE t2
SET t2.ry = t1.ry
from temptablell t2
JOIN temptablell t1 
ON t1.row = (t2.row - 1) AND T1.prono=T2.prono
WHERE t1.sochuapc = 'A'
AND t2.sochuapc = 'B'
;
------------//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////// run
delete temptablell where ry is null;
--update id trong bang temptablell
WITH NumberedRows AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY prono ORDER BY prono) AS RowNum
    FROM temptablell
)
UPDATE NumberedRows
SET ID = (RowNum + 1) / 2;
--xoa ry du thua do tao lap nhieu bang
delete temptableas2 where ry is null;
-------------------------------//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

--update size ----------------------------------------------------------------------------
-------------------------------------------------------------------

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size1), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 3.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size2), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 3.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size3), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 4.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size4), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 4.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size5), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 5.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size6), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 5.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size7), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 6.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size8), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 6.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size9), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 7.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size10), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 7.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size11), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 8.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size12), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 8.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size13), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 9.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size14), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 9.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size15), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 10.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size16), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 10.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size17), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 11.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size18), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 11.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size19), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 12.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size20), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 12.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size21), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 13.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size22), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 13.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size23), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 14.0;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size24), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 14.5;

UPDATE temptableas2
SET qty = (
    SELECT COALESCE(SUM(temptablell.size25), 0) 
    FROM temptablell
    WHERE temptablell.prono = temptableas2.prono 
    AND temptablell.id = temptableas2.id
    GROUP BY temptablell.id
)
WHERE temptableas2.size = 15.0;
---
---			
--xoa nhung size co qty la null
Delete from temptableas2
where qty is null
;
UPDATE temptableas2
SET seq = RIGHT(temptablell.Ry, 2)
FROM temptableas2
JOIN temptablell ON temptablell.prono = temptableas2.prono and temptablell.id = temptableas2.id
WHERE temptablell.Ry LIKE '%TUA%';
;
Update temptableas2
set seq=1
from temptableas2
where seq is null

;
-------------------------------------------------------------------------------------------------------------------------
UPDATE temptableas2
SET ry = LEFT(ry, CHARINDEX('TUA', ry) - 2)
WHERE ry LIKE '%TUA%';
;
--xoa trung lap
WITH cte AS (
  SELECT prono,seq,dangfom, ry, size, ROW_NUMBER() OVER (PARTITION BY prono,seq,dangfom, ry, size ORDER BY (SELECT NULL)) AS rn
  FROM temptableas2
)
DELETE FROM cte
WHERE rn > 1;
;
--
update A1
set A1.dangfom=A2.dangfom,
	A1.goo=A2.goo,
	A1.may=A2.may
 from 
(select * from temptablell 
where convert(varchar,id)+prono in(select distinct convert(varchar,a.id)+a.Prono
from temptablell a
join temptablell b
on a.prono=b.prono and a.id=b.id
where a.sochuapc = 'A' and a.dangfom is null ) and sochuapc='A')A1
join (select * from temptablell)A2 on A1.prono=A2.prono and A2.id+1=A1.id;
--
UPDATE t2
SET t2.dangfom = t1.dangfom,
	t2.goo = t1.goo,
	t2.may = t1.may
from temptablell t2
JOIN temptablell t1 
ON t1.row = (t2.row - 1) AND T1.prono=T2.prono
WHERE t1.sochuapc = 'A'
AND t2.sochuapc = 'B';
--
update temptableas2
set dangfom =temptablell.dangfom,
	goo =temptablell.goo,
	may =temptablell.may
from temptableas2
join temptablell on temptablell.prono=temptableas2.prono and temptablell.id=temptableas2.id;
--
delete temptableas2
where qty = 0;
--
ALTER TABLE temptableas2
ALTER COLUMN size VARCHAR(20);
--
UPDATE temptableas2
SET size = 
    CASE 
        WHEN CHARINDEX('.', CAST(size AS VARCHAR(20))) = 2 THEN '0' + CAST(size AS VARCHAR(20))
        WHEN CHARINDEX('.', CAST(size AS VARCHAR(20))) = 3 THEN CAST(size AS VARCHAR(20))
    END
WHERE size IS NOT NULL; -- Điều kiện WHERE để chỉ cập nhật các hàng có giá trị size khác NULL
--
--Xóa khoảnh trắng ở cuối ry
UPDATE temptableas2
SET ry = RTRIM(ry);
--xóa những hàng trùng size,ry,qty,dangfom giữ lại 1 prono
WITH CTE AS (
    SELECT ry, seq, dangfom, size, qty, prono,
           ROW_NUMBER() OVER (PARTITION BY ry, seq, dangfom, size, qty ORDER BY prono) AS RowNum
    FROM temptableas2
)
DELETE FROM CTE
WHERE RowNum > 1;
--
SELECT DISTINCT t2.ry, t2.seq
FROM temptableas2 t2
LEFT JOIN Timestamp ts ON ts.TS_Tour = t2.dangfom
WHERE ts.TS_Tour IS NULL;

Drop table temptablell
;
Drop table temptableas2
;
drop table #TempData5
;
drop table #TempData22
;";

            try
            {
                DataTable resultTable = dbConnect.ExecuteQuery(sql);

                if (resultTable.Rows.Count > 0)
                {
                    //Nếu lỗi thì xóa bang xoay tua để chạy procedure không bị lỗi
                    dbConnect.ExecuteQuery(@"delete BANG_XOAY_TUA");

                    //string message = "Thông báo: Có vấn đề về size ở các lệnh:\r\n";
                    string message = "";

                    // Duyệt qua từng hàng và lấy giá trị từ các cột
                    foreach (DataRow row in resultTable.Rows)
                    {
                        // Giả sử bạn muốn hiển thị giá trị của cột đầu tiên
                        // Bạn có thể thay đổi theo yêu cầu cụ thể của bạn
                        message += row[0].ToString() + " TUA ";
                        message += row[1].ToString() + "\r\n";
                    }

                    // Tạo và cấu hình một TextBox để hiển thị thông báo
                    TextBox textBox = new TextBox
                    {
                        Text = message,
                        Multiline = true,
                        ReadOnly = true,
                        Dock = DockStyle.Fill,
                        ScrollBars = ScrollBars.Both,
                        WordWrap = true
                    };

                    // Tạo một Form để chứa TextBox
                    Form form = new Form
                    {
                        Text = "Không tìm thấy giờ ở các lệnh, tua sau:",
                        Width = 600,
                        Height = 400,
                        StartPosition = FormStartPosition.CenterScreen
                    };

                    // Thêm TextBox vào Form
                    form.Controls.Add(textBox);

                    // Đăng ký sự kiện Click để chọn tất cả văn bản trong TextBox
                    //textBox.Click += (sender, e) => textBox.SelectAll();

                    // Hiển thị Form
                    form.ShowDialog();

                    // Thoát ứng dụng sau khi form được đóng
                    Application.Exit();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi khi thực hiện kiểm tra size: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        //
    }
}

