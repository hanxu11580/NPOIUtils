using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
public static class NPOIUtils
{
    public static ExcelStrategy DefaultExcelStategy = new ExcelStrategy();

    /// <summary>
    /// 创建Excel
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="sheetNames"></param>
    public static void CreateExcel(string filePath, string sheetName)
    {
        if (File.Exists(filePath))
        {
            throw new Exception($"FilePath:{filePath} is Exist...");
        }
        IWorkbook workbook = CheckAndCreateIWorkBook(filePath);
        ISheet sheet = workbook.CreateSheet(sheetName);
        DefaultExcelStategy.SetBlankExcelLayout(sheet);
        using (var fs = new FileStream(filePath, FileMode.Create))
        {
            workbook.Write(fs);
        }

        workbook.Close();
    }

    public static IWorkbook ReadWorkBook(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new Exception($"FilePath:{filePath} is NonExist...");
        }

        FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        IWorkbook workbook = CheckAndReadIWorkBook(filePath, fs);

        return workbook;
    }

    /// <summary>
    /// 更新
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="mainKey"></param>
    /// <param name="fieldName"></param>
    /// <param name="val"></param>
    public static void UpdateSheet(ISheet sheet, string mainKey, string fieldName, object val)
    {
        IWorkbook workbook = sheet.Workbook;
        DefaultExcelStategy.UpdateSheet(sheet, mainKey, fieldName, val);
    }


    #region helper

    private static IWorkbook CheckAndCreateIWorkBook(string filePath)
    {
        IWorkbook workbook = null;
        string extensionName = Path.GetExtension(filePath);
        if (extensionName == ".xls")
        {
            workbook = new HSSFWorkbook();
        }
        else if (extensionName == ".xlsx")
        {
            workbook = new XSSFWorkbook();
        }
        else
        {
            throw new Exception("ExtensionName Error...");
        }

        return workbook;
    }


    private static IWorkbook CheckAndReadIWorkBook(string filePath, Stream stream)
    {
        IWorkbook workbook = null;
        string extensionName = Path.GetExtension(filePath);
        if (extensionName == ".xls")
        {
            workbook = new HSSFWorkbook(stream);
        }
        else if (extensionName == ".xlsx")
        {
            workbook = new XSSFWorkbook(stream);
        }
        else
        {
            throw new Exception("ExtensionName Error...");
        }

        return workbook;
    }

    #endregion
}

