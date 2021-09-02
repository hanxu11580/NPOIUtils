using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

public class ExcelStrategy: IExcelStrategy
{
    // 默认id为主键

    // 行 一行空格 + 3行
    public int rowStart = 4;
    // 列 一列空格 + 1列
    public int columnStart = 2;
    // 列字段描述列表
    public List<ColumnDesc> columnDescs;

    public ExcelStrategy()
    {
        columnDescs = new List<ColumnDesc>();
    }

    public void SetBlankExcelLayout(ISheet sheet)
    {
        CellSytleHelper.SetWorkbook(sheet.Workbook);
        IRow r1 = sheet.CreateRow(1);
        ICell r1Cell = r1.CreateCell(1);
        r1Cell.SetCellValue("字段名");
        r1Cell.CellStyle = CellSytleHelper.style_Green;
        IRow r2 = sheet.CreateRow(2);
        ICell r2Cell = r2.CreateCell(1);
        r2Cell.SetCellValue("字段描述");
        r2Cell.CellStyle = CellSytleHelper.style_Yellow;
        IRow r3 = sheet.CreateRow(3);
        ICell r3Cell = r3.CreateCell(1);
        r3Cell.SetCellValue("字段类型");
        r3Cell.CellStyle = CellSytleHelper.style_Bule;
    }

    public void UpdateSheet(ISheet sheet, string mainKey, string fieldName, object val)
    {
        IRow r1 = sheet.GetRow(1);
        IRow r2 = sheet.GetRow(2);
        IRow r3 = sheet.GetRow(3);

        for (int i = columnStart; i < r1.LastCellNum; i++)
        {
            columnDescs.Add(new ColumnDesc()
            {
                fieldColumnIdx = i,
                fieldName = r1.GetCell(i).StringCellValue,
                fieldDesc = r2.GetCell(i).StringCellValue,
                fieldType = r3.GetCell(i).StringCellValue,
            });
        }

        // 先找到主键相同的行
        for (int i = rowStart; i < sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);
            if (mainKey.Equals(row.GetCell(columnStart).StringCellValue, StringComparison.Ordinal))
            {
                var columnDesc = columnDescs.First(cd => cd.fieldName == fieldName);
                var idx = columnDesc.fieldColumnIdx;
                if(columnDesc.fieldType == "int")
                {
                    row.GetCell(idx).SetCellValue((int)val);
                }else if(columnDesc.fieldType == "bool")
                {
                    row.GetCell(idx).SetCellValue((bool)val);
                }
                else
                {
                    row.GetCell(idx).SetCellValue((string)val);
                }
                break;
            }
        }
    }
}

public class ColumnDesc
{
    // 这个字段在哪一列
    public int fieldColumnIdx;
    public string fieldName;
    public string fieldDesc;
    public string fieldType;
}


public static class CellSytleHelper
{
    public static ICellStyle style_Green;
    public static ICellStyle style_Yellow;
    public static ICellStyle style_Bule;

    public static void SetWorkbook(IWorkbook workbook)
    {
        style_Green = workbook.CreateCellStyle();
        style_Green.FillPattern = FillPattern.SolidForeground;
        style_Green.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
        style_Yellow = workbook.CreateCellStyle();
        style_Yellow.FillPattern = FillPattern.SolidForeground;
        style_Yellow.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightYellow.Index;
        style_Bule = workbook.CreateCellStyle();
        style_Bule.FillPattern = FillPattern.SolidForeground;
        style_Bule.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightTurquoise.Index;
    }
}
