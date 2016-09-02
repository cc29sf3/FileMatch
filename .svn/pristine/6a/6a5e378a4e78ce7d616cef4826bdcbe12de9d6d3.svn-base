using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Manual_Import.Model;
using System.IO;
using Utility.Dao;
using System.Data;

namespace Manual_Import
{
    /// <summary>
    /// FileTypeDel.xaml 的交互逻辑
    /// </summary>
    public partial class FileTypeDel : Window
    {
        public FileTypeDel()
        {
            InitializeComponent();
        }
        public FileTypeDel(SQLiteDBHelper db,string taskType)
        {
            InitializeComponent();
            this.db = db;
            this.TaskType = taskType;
        }

        private SQLiteDBHelper db;
        private string TaskType;

        private void Btn_Del_Click(object sender, RoutedEventArgs e)
        {
            TabItem ti = GetSelectTab();
            if (ti == null)
                return;
            MessageBoxResult mbr = MessageBox.Show("确定删除所有"+ti.Header+"文件吗?", "警告", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (mbr == MessageBoxResult.OK)
            {
                DataGrid grid = ti.Content as DataGrid;
                foreach (var item in grid.Items)
                {
                    DelFile df = item as DelFile;
                    File.Delete( df.FullPath +"\\" +df.Name);
                    Db_Delte_Sync(df.FullPath + "\\" + df.Name);
                }
                grid.ItemsSource = null;
                MessageBox.Show("删除成功!");
            }
            
        }

        private TabItem GetSelectTab()
        {
            foreach (TabItem ti in Tab_File.Items)
            {
                if (ti.IsSelected)
                {
                    return ti;
                }
            }
            return null;
        }

        private void Db_Delte_Sync(string fullpath)
        {
            DataTable dt = db.ExecuteDataTable("select 编号 from db_File where 编号=(select 编号 from db_File where 路径='" + fullpath + "')", null);
            if (dt.Rows.Count == 1)//如果该任务只包含这一个文件,则要把对应任务也删除
            {
                string code = dt.Rows[0][0].ToString();
                db.ExecuteNonQuery("delete from XW_FileOrderinfo where 编号='" + code + "'", null);
                db.ExecuteNonQuery("delete from db_File where 路径 like '" + fullpath + "%'", null);
                if (TaskType == "刊盘")
                    db.ExecuteNonQuery("delete from db_State where 编号='" + code + "'", null);
            }
            else if (dt.Rows.Count > 1)
            {
                db.ExecuteNonQuery("delete from db_File where 路径 like '" + fullpath + "%'", null);
            }
        }
    }
}
