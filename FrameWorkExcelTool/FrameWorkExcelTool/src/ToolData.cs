using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public static class ToolData
{
    private static RecordInfo record;

    public static RecordInfo Record
    {
        get {

            JudgeInit();
            return record;

        }
        set
        {
            record = value;
            record.Change();
            RecordManager.SaveRecord("Record", "record", record);
        }
    }

    static bool isInit = false;
    static void JudgeInit()
    {
        if(!isInit)
        {
            isInit = true;

            record = RecordManager.GetRecord("Record", "record", new RecordInfo());
        }
    }
}

public class RecordInfo : INotifyCollectionChanged
{
    public string excelPath;
    public string exportPath;

    public List<CheckInfo> checkList = new List<CheckInfo>();

    public string ExcelPath { get => excelPath; set => excelPath = value; }
    public string ExportPath { get => exportPath; set => exportPath = value; }
    public List<CheckInfo> CheckList
    {
        get
        {
            if (!string.IsNullOrEmpty(excelPath))
            {
                List<string> list = FileTool.GetAllFileNamesByPath(excelPath,new string[] { "xls", "xlsx", "csv" },true);

                for (int i = 0; i < list.Count; i++)
                {
                    if(GetCheckInfo(list[i]) == null)
                    {
                        CheckInfo info = new CheckInfo();

                        string fileName = FileTool.RemoveExpandName( FileTool.GetFileNameByPath(list[i]));

                        info.FilePath = list[i];
                        info.FileName = fileName;
                        info.ExportName = fileName;

                        checkList.Add(info);
                    }
                }

                //移除不存在的文件
                for (int i = 0; i < checkList.Count; i++)
                {
                    if(!list.Contains(checkList[i].FilePath))
                    {
                        checkList.RemoveAt(i);
                        i--;
                    }
                }
            }

            return checkList;
        }
        set
        {
            checkList = value;
            ToolData.Record = ToolData.Record;
        }
    }

    public event NotifyCollectionChangedEventHandler CollectionChanged;

    public void Change()
    {
        NotifyCollectionChangedEventArgs e = new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset);
        CollectionChanged?.Invoke(this, e);
    }

    public CheckInfo GetCheckInfo(string filePath)
    {
        for (int i = 0; i < checkList.Count; i++)
        {
            if(checkList[i].FilePath == filePath)
            {
                return checkList[i];
            }
        }

        return null;
    }
}

public class CheckInfo
{
    public string fileName;
    public string exportName;
    public bool isCheck;
    public string filePath;

    public bool IsCheck { get => isCheck; set => isCheck = value; }
    public string FileName { get => fileName; set => fileName = value; }
    public string FilePath { get => filePath; set => filePath = value; }
    public string ExportName { get => exportName; set => exportName = value; }
}
