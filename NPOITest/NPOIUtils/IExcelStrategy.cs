using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

interface IExcelStrategy
{
    /// <summary>
    /// 设置空白表格式
    /// </summary>
    /// <param name="sheet"></param>
    void SetBlankExcelLayout(ISheet sheet);

    void UpdateSheet(ISheet sheet, string mainKey, string fieldName, object val);
}
