// NPOI lib
using NPOI.HSSF.UserModel; // XSSF for xls
using NPOI.SS.UserModel;   // SS for generic model (guess)
using NPOI.XSSF.UserModel; // XSSF for xlsx


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApplication3
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            testExcel();
        }

        public class ExcelCell
        {
            public XSSFCell cell;
            public ExcelCell(XSSFCell c) { cell = c; }

            public int row { get { return cell.RowIndex; } }
            public int col { get { return cell.ColumnIndex; } }
            public XSSFCellStyle cellStyle { get { return (XSSFCellStyle)cell.CellStyle; } }
            public BorderStyle bsTop    { get { return cellStyle.BorderTop; } }
            public BorderStyle bsBottom { get { return cellStyle.BorderBottom; } }
            public BorderStyle bsLeft   { get { return cellStyle.BorderLeft; } }
            public BorderStyle bsRight  { get { return cellStyle.BorderRight; } }
            public XSSFColor fgColor { get { return (XSSFColor)cellStyle.FillForegroundColorColor; } }
        }
        private void testExcel()
        {
            List<ExcelCell> ExcelTable = new List<ExcelCell>();
            string strFilePath = string.Format(@"C:\1\Book1.xlsx");
            XSSFWorkbook workbook; // xlsx, HSSF for xls
            Stopwatch sw = new Stopwatch();

            sw.Start();
            using (FileStream fs = File.Open(strFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                workbook = new XSSFWorkbook(fs);
                XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);
                //sheet = workbook.GetSheet(string name);
                int RowNum = sheet.PhysicalNumberOfRows;

                for (int i = 0; i < RowNum; ++i)
                {
                    XSSFRow row = (XSSFRow)sheet.GetRow(i);
                    //ICell cell = sheet.GetRow(i).GetCell(j);
                    foreach (XSSFCell cell in row.Cells)
                        ExcelTable.Add(new ExcelCell(cell));
                }
            }
            sw.Stop();

            //XSSFFont font = ExcelTable[1].cellStyle.GetFont();
            MessageBox.Show(sw.ElapsedMilliseconds.ToString());
        }
    }
}
