using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console_20240210_001
{
    public static class ExcelUtil
    {
        public static void ExportDataTableToExcel(DataTable table, string filePath, string[]? selectedColumns)
        {
            // 确保EPPlus不处于商业许可模式
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if(selectedColumns==null)
            {
                IList<string> columns =new List<string>();
                foreach(DataColumn dc in table.Columns)
                {
                    columns.Add(dc.ColumnName);
                }

                selectedColumns=columns.ToArray();
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // 添加选定的列头
                for (int i = 0; i < selectedColumns.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = selectedColumns[i];
                }

                // 填充数据行
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    for (int j = 0; j < selectedColumns.Length; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = table.Rows[i][selectedColumns[j]];
                    }
                }

                // 保存到文件
                FileInfo fileInfo = new FileInfo(filePath);
                if (fileInfo.Exists)fileInfo.Delete();
                package.SaveAs(fileInfo);
            }
        }
    }
}
