using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace FrameWorkExcelTool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Grid_main.DataContext = ToolData.Record;
        }

        private void Button_ClickExcelPath(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog m_Dialog = new FolderBrowserDialog();
            DialogResult result = m_Dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            string m_Dir = m_Dialog.SelectedPath.Trim();

            this.Text_ExcelPath.Text = m_Dir;

            ToolData.Record.ExcelPath = m_Dir;

            ToolData.Record.CheckList = ToolData.Record.CheckList;
            ToolData.Record = ToolData.Record;

            List_ExcelList.ItemsSource = ToolData.Record.CheckList;
        }

        private void Button_ClickExportPath(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog m_Dialog = new FolderBrowserDialog();
            DialogResult result = m_Dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            string m_Dir = m_Dialog.SelectedPath.Trim();

            this.Text_ExportPath.Text = m_Dir;


            ToolData.Record.ExportPath = m_Dir;
            ToolData.Record = ToolData.Record;

            
        }

        private void Button_ClickExport(object sender, RoutedEventArgs e)
        {
            try
            {
                List<CheckInfo> list = ToolData.Record.CheckList;

                for (int i = 0; i < list.Count; i++)
                {
                    if (list[i].IsCheck)
                    {
                        ExcelTool.ExportExcel(list[i].FilePath, ToolData.Record.ExportPath + "\\" + list[i].ExportName + ".txt");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString());
            }
            finally
            {
                System.Windows.MessageBox.Show("导出完成");
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.CheckBox cb = sender as System.Windows.Controls.CheckBox;

            string fileName = (string)cb.Tag;

            CheckInfo ci = ToolData.Record.GetCheckInfo(fileName);

            if(ci != null)
            {
                ci.isCheck = cb.IsChecked ?? false;

                ToolData.Record.CheckList = ToolData.Record.CheckList;
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox cb = sender as System.Windows.Controls.TextBox;
            string fileName = (string)cb.Tag;

            CheckInfo ci = ToolData.Record.GetCheckInfo(fileName);
            if (ci != null)
            {
                ci.exportName = cb.Text;
                ToolData.Record.CheckList = ToolData.Record.CheckList;
            }
        }

        private object ExitFrames(object frame)
        {
            ((DispatcherFrame)frame).Continue = false;
            return null;
        }
    }
}
