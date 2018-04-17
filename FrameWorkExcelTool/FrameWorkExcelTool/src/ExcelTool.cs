using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class ExcelTool
{
    public static void ExportExcel(string excelPath, string exportPath)
    {
        FileTool.CreatFilePath(exportPath);

        ExcelReader excelReader = new ExcelReader();
        DataTable dataTable = excelReader.ReadFile(excelPath);
        StreamWriter streamWriter = File.CreateText(exportPath);

        int rows = dataTable.Rows.Count;
        int columns = dataTable.Columns.Count;

        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                string content = dataTable.Rows[i][j].ToString();
                bool isSpace = j < columns - 1;
                if (isSpace)
                {
                    content += "\t";
                }
                streamWriter.Write(content);
            }
            bool isLine = i < rows - 1;
            if (isLine)
            {
                streamWriter.WriteLine("");
            }
        }
        streamWriter.Close();
    }
}
