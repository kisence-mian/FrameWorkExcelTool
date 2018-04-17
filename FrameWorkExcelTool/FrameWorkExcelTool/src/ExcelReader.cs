using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class ExcelReader
{
    private enum FileType
    {
        noset,
        xls,
        xlsx,
        csv
    }
    private string filePath;
    private string fileName;
    private OleDbConnection conn;
    private DataTable readDataTable;
    private string connString;
    private ExcelReader.FileType fileType = ExcelReader.FileType.noset;
    private void SetFileInfo(string path)
    {
        this.filePath = path;
        this.fileName = this.filePath.Remove(0, this.filePath.LastIndexOf("\\") + 1);
        string expandName = this.fileName.Split(new char[]
        {
                '.'
        })[1];
        if (!(expandName == "xls"))
        {
            if (!(expandName == "xlsx"))
            {
                if (expandName == "csv")
                {
                    this.connString = string.Concat(new string[]
                    {
                            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=",
                            this.filePath.Remove(this.filePath.LastIndexOf("\\") + 1),
                            ";Extended Properties='Text;FMT=Delimited;HDR=",
                            "NO",
                            ";'"
                    });
                    this.fileType = ExcelReader.FileType.csv;
                }
            }
            else
            {
                this.connString = string.Concat(new string[]
                {
                        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=",
                        this.filePath,
                        ";Extended Properties='Excel 12.0;HDR=",
                        "NO",
                        "; IMEX=1;'"
                });
                this.fileType = ExcelReader.FileType.xlsx;
            }
        }
        else
        {
            this.connString = string.Concat(new string[]
            {
                    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=",
                    this.filePath,
                    ";Extended Properties='Excel 8.0;HDR=",
                    "NO",
                    ";IMEX=1;'"
            });
            this.fileType = ExcelReader.FileType.xls;
        }
    }
    public DataTable ReadFile(string path)
    {
        bool flag = File.Exists(path);
        if (flag)
        {
            this.SetFileInfo(path);
            using (this.conn = new OleDbConnection(this.connString))
            {
                this.conn.Open();
                DataTable oleDbSchemaTable = this.conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string text = (this.fileType == ExcelReader.FileType.csv) ? this.fileName : oleDbSchemaTable.Rows[0][2].ToString().Trim();
                string selectCommandText = string.Empty;
                selectCommandText = "Select   *   From   [" + text + "]";
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(selectCommandText, this.conn);
                DataSet dataSet = new DataSet();
                oleDbDataAdapter.Fill(dataSet, text);
                this.readDataTable = dataSet.Tables[0];
            }
        }
        return this.readDataTable;
    }
}