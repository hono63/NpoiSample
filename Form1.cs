using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace NpoiSample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            CreateSample();

            /*try
            {
                string filePath = "sample.xlsx";

                //ブック作成
                var book = CreateNewBook(filePath);

                //シート無しのexcelファイルは保存は出来るが、開くとエラーが発生する
                book.CreateSheet("newSheet");

                //ブックを保存
                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    book.Write(fs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }*/
        }


        //ブック作成
        private IWorkbook CreateNewBook(string filePath)
        {
            IWorkbook book;
            var extension = Path.GetExtension(filePath);

            // HSSF => Microsoft Excel(xls形式)(excel 97-2003)
            // XSSF => Office Open XML Workbook形式(xlsx形式)(excel 2007以降)
            if (extension == ".xls")
            {
                book = new HSSFWorkbook();
            }
            else if (extension == ".xlsx")
            {
                book = new XSSFWorkbook();
            }
            else
            {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            return book;
        }

        private void CreateSample()
        {
            try
            {
                //ブック読み込み
                var book = WorkbookFactory.Create("sample.xlsx");

                //シート名からシート取得
                var sheet = book.GetSheet("newSheet");

                //セルに設定
                WriteCell(sheet, 0, 0, "0-0");
                WriteCell(sheet, 1, 1, "1-1");
                WriteCell(sheet, 0, 3, 100);
                WriteCell(sheet, 0, 4, DateTime.Today);

                //日付表示するために書式変更
                var style = book.CreateCellStyle();
                style.DataFormat = book.CreateDataFormat().GetFormat("yyyy/mm/dd");
                WriteStyle(sheet, 0, 4, style);

                //ブックを保存
                using (var fs = new FileStream("sample2.xlsx", FileMode.Create))
                {
                    book.Write(fs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        //セル設定(文字列用)
        private void WriteCell(ISheet sheet, int columnIndex, int rowIndex, string value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        //セル設定(数値用)
        private void WriteCell(ISheet sheet, int columnIndex, int rowIndex, double value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        //セル設定(日付用)
        private void WriteCell(ISheet sheet, int columnIndex, int rowIndex, DateTime value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
        }

        //書式変更
        private void WriteStyle(ISheet sheet, int columnIndex, int rowIndex, ICellStyle style)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.CellStyle = style;
        }
    }
}
