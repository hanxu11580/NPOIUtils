using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace NPOITest
{
    class Program
    {
        static void Main(string[] args)
        {
            string filepath = "../../../ExcelFiles/abc.xls";

            //NPOIUtils.CreateExcel(filepath, "表1");

            // 修改
            IWorkbook workbook = NPOIUtils.ReadWorkBook(filepath);
            ISheet sheet = workbook.GetSheet("表1");

            NPOIUtils.UpdateSheet(sheet, "id2", "name", "hanx");
            NPOIUtils.UpdateSheet(sheet, "id3", "age", 500);
            using (var fs = new FileStream(filepath, FileMode.Open, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            // 读取所有数据
            //for (int i = startRow+1; i < sheet.LastRowNum; i++)
            //{
            //    var row = sheet.GetRow(i);
            //    if (row != null)
            //    {
            //        string rowStr = string.Empty;
            //        for (int j = startCell+1; j < row.LastCellNum; j++)
            //        {
            //            ICell cell = row.GetCell(j);
            //            rowStr += cell.ToString();
            //        }
            //        Console.WriteLine(rowStr);
            //    }
            //}

            // 读某一条数据


            Console.ReadKey();
        }
    }
}
