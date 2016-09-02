﻿using System;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Manual_Import.ViewModel;
using Manual_Import.Model;
using Manual_Import;
using Manual_Import.Helper;
using System.Xml.Linq;
using System.Threading.Tasks;
using System.Data.SQLite;
using FileMatch;
using Utility.Common;
using FileMatch.Entity;
using SharpCompress.Reader;
using SharpCompress.Common;
using System.Data;
using System.Net;
using Utility.Log;
using Utility.Dao;
using Utility.Submit;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;
using System.Collections.ObjectModel;


namespace Manual_Import
{
    
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        ViewModel_Main ViewModel;
        TidyHelper tHelper;
        SQLiteDBHelper db;
        Action<string> Render;
       
        bool Button4Checked=false;
        Dictionary<string, string> TempDic=new Dictionary<string,string>();//保存每个篇任务是否可做
        //ListViewSimpleAdorner myAdorner;//鼠标画框层
        Point? myDragStartPoint;//画框时鼠标的起始坐标
        CancellationTokenSource tokenSource ;

        bool IS_CUSTOM_DEFINED=false;//是否用户自定义提取页数

        public MainWindow()
        {
            InitializeComponent();
            //myAdorner = new ListViewSimpleAdorner(View_Work);
            Render = RenderView;
           
        }

        public void ResetView()
        {
            this.Dispatcher.Invoke(new Action(() => {
                this.ViewModel.Models.Clear();
                this.ViewModel.TaskCode = "";
                this.ViewModel.TaskType = "";
                this.ViewModel = null;
                radio_muti.IsChecked = true;
                radio_only.IsChecked = false;
                GetTask();
            }));
        }
        /// <summary>
        /// 纯盘任务:查询文件状态
        /// </summary>
        /// <returns>未整理:0 整理成功:1 整理失败:-1 篇提交:2 文件夹:-2</returns>
        public int GetFileState(string filePath)
        {
            string sql="";
            if (ViewModel.TaskType == "纯盘")
                sql = "select a.可读,a.提取,b.提交否 from db_File a join XW_FileOrderinfo b on a.编号=b.编号   where a.路径=\"" + filePath + "\"";
            else
                sql = "select a.可读,a.提取,b.保存否 from db_File a join db_State b on a.编号=b.编号 where a.路径=\"" + filePath + "\"";
            DataTable dt = db.ExecuteDataTable(sql, null);
            if (dt.Rows.Count == 0)
                return 0;
            else
            {
                DataRow row = dt.Rows[0];
                if (row[2].ToString() == "是")
                    return 2;
                else
                {
                    if (row[0].ToString() != "是" || row[1].ToString() != "是")
                        return -1;
                    else
                        return 1;
                }
            }
        }


      
        #region 渲染列表
     

        public void RenderView(string path)
        {
            DirectoryInfo rootDir = new DirectoryInfo(path);
            ViewModel.CurPath = path;
            ViewModel.Models.Clear();
            //ObservableCollection<Model_FileSystem> PreModels = new ObservableCollection<Model_FileSystem>();
            var fileCollection = rootDir.GetFileSystemInfos();
            progressBar.Maximum = fileCollection.Count();
            progressBar.Visibility = Visibility.Visible;
            Task task = new Task(new Action(() =>
            {
                try
                {
                    foreach (FileSystemInfo fsi in fileCollection)
                    {
                        Model_FileSystem model = new Model_FileSystem() { Name = fsi.Name, FullPath = fsi.FullName, Time = fsi.LastWriteTime.ToString("yyyy/MM/dd hh:mm") };
                        if (fsi.Attributes == FileAttributes.Directory)
                        {
                            model.Type = SystemType.Dir;
                            model.Extension = "文件夹";
                            if (Traverse(fsi.FullName) == Convert.ToInt32(db.ExecuteScalar("select count(1) from db_File where 路径 like '" + fsi.FullName + "\\%'", null)))
                                model.HasTidy = 1;
                            else
                                model.HasTidy = -2;
                            //PreModels.Insert(0, model);
                            this.Dispatcher.Invoke(new Action(() => { ViewModel.Models.Insert(0, model); }));
                        }
                        else
                        {
                            try
                            {
                                model.FileSize = ((FileInfo)fsi).Length / 1024 + "KB";
                                model.HasTidy = GetFileState(model.FullPath);
                                if (fsi.Extension.ToLower() == ".doc" || fsi.Extension.ToLower() == ".docx" || fsi.Extension.ToLower() == ".wps")
                                    model.Type = SystemType.Word;
                                else if (fsi.Extension.ToLower() == ".pdf")
                                    model.Type = SystemType.PDF;
                                else
                                    model.Type = SystemType.File;
                                model.Extension = fsi.Extension.Substring(1).ToUpper() + "文件";
                            }
                            catch (Exception e)
                            {
                                //model.HasTidy = 0;
                            }
                            //PreModels.Add(model);
                            this.Dispatcher.Invoke(new Action(() => { ViewModel.Models.Add(model); }));
                            
                        }
                        this.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            progressBar.Value++;
                        }));
                    }
                }
                catch (Exception eeee)
                {
                    ShowLog(eeee.Message);
                    TextLog.WritwLog(eeee.Message);
                }
            }));
            task.ContinueWith(t =>
            {
                this.Dispatcher.BeginInvoke(new Action(() => {
                    //ViewModel.Models = PreModels;
                    if (ViewModel.RootPath != path)
                    {
                        Model_FileSystem model = new Model_FileSystem() { Name = "..", FullPath = path.Substring(0, path.LastIndexOf('\\')) };
                        ViewModel.Models.Insert(0, model);
                    }
                    progressBar.Visibility = Visibility.Collapsed;
                    progressBar.Value= 0;
                    
                }));
            });
            task.Start();
        }

        /// <summary>
        /// 整理成功的文件,加载文件列表
        /// </summary>
        /// <param name="path"></param>
        public void RenderView_HasTidy(string path)
        {
            DirectoryInfo rootDir = new DirectoryInfo(path);
            ViewModel.CurPath = path;
            ViewModel.Models.Clear();
            //ObservableCollection<Model_FileSystem> PreModels = new ObservableCollection<Model_FileSystem>();
            var fileCollection = rootDir.GetFileSystemInfos();
            progressBar.Maximum = fileCollection.Count();
            progressBar.Visibility = Visibility.Visible;
            Task task = new Task(() => {
                foreach (FileSystemInfo fsi in rootDir.GetFileSystemInfos())
                {
                    Model_FileSystem model = new Model_FileSystem() { Name = fsi.Name, FullPath = fsi.FullName };
                    if (fsi.Attributes == FileAttributes.Directory)
                    {
                        model.Type = SystemType.Dir;
                        try
                        {
                            int count = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 路径 like '" + model.FullPath + "\\%' and (提取='是' and 可读='是')", null));
                            if (count != 0)
                            {
                                model.HasTidy = -2;
                                //PreModels.Insert(0, model);
                                this.Dispatcher.Invoke(new Action(() => { ViewModel.Models.Insert(0, model); }));

                            }
                        }
                        catch (Exception e)
                        {
                        }
                    }
                    else
                    {
                        if (fsi.Extension.ToLower() == ".doc" || fsi.Extension.ToLower() == ".docx" || fsi.Extension.ToLower() == ".wps")
                            model.Type = SystemType.Word;
                        else if (fsi.Extension.ToLower() == ".pdf")
                            model.Type = SystemType.PDF;
                        else
                            model.Type = SystemType.File;
                        try
                        {
                            SQLiteDataReader sqliteReader = db.ExecuteReader("select * from db_File where 路径='" + model.FullPath + "' and (提取='是' and 可读='是')", null);
                            if (sqliteReader.HasRows)
                            {
                                model.HasTidy = 1;
                                //PreModels.Add(model);
                                this.Dispatcher.Invoke(new Action(() => { ViewModel.Models.Add(model); }));
                            }
                        }
                        catch (Exception e)
                        {
                        }
                    }
                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        progressBar.Value++;
                    }));
                }
            });
            task.ContinueWith(t => {
                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    //ViewModel.Models = PreModels;
                    if (ViewModel.RootPath != path)
                    {
                        Model_FileSystem model = new Model_FileSystem() { Name = "..", FullPath = path.Substring(0, path.LastIndexOf('\\')) };
                        ViewModel.Models.Insert(0, model);
                    }
                    progressBar.Visibility = Visibility.Collapsed;
                    progressBar.Value = 0;
                }));
            });
            task.Start();
        }

        /// <summary>
        /// 未整理过的文件加载文件列表
        /// </summary>
        /// <param name="path"></param>
        public void RenderView_UnTidy(string path)
        {
            DirectoryInfo rootDir = new DirectoryInfo(path);
            ViewModel.CurPath = path;
            ViewModel.Models.Clear();
           // ObservableCollection<Model_FileSystem> PreModels = new ObservableCollection<Model_FileSystem>();
            var fileCollection = rootDir.GetFileSystemInfos();
            progressBar.Maximum = fileCollection.Count();
            progressBar.Visibility = Visibility.Visible;
            Task task = new Task(new Action(() =>
            {
                foreach (FileSystemInfo fsi in fileCollection)
                {
                    Model_FileSystem model = new Model_FileSystem() { Name = fsi.Name, FullPath = fsi.FullName };
                    if (fsi.Attributes == FileAttributes.Directory)
                    {
                        model.Type = SystemType.Dir;
                        try
                        {
                            int count = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 路径 like '" + model.FullPath + "\\%'", null));
                            int fs = Traverse(model.FullPath);
                            if (count < fs)
                            {
                                model.HasTidy = -2;
                                //PreModels.Insert(0, model);
                                this.Dispatcher.Invoke(new Action(() => { ViewModel.Models.Insert(0, model); }));
                            }
                        }
                        catch (Exception e)
                        {
                            //PreModels.Insert(0, model);
                            this.Dispatcher.Invoke(new Action(() => { ViewModel.Models.Insert(0, model); }));
                        }
                    }
                    else
                    {
                        if (fsi.Extension.ToLower() == ".doc" || fsi.Extension.ToLower() == ".docx" || fsi.Extension.ToLower() == ".wps")
                            model.Type = SystemType.Word;
                        else if (fsi.Extension.ToLower() == ".pdf")
                            model.Type = SystemType.PDF;
                        else
                            model.Type = SystemType.File;
                        try
                        {
                            SQLiteDataReader sqliteReader = db.ExecuteReader("select * from db_File where 路径='" + model.FullPath + "'", null);
                            if (!sqliteReader.HasRows)
                            {
                                model.HasTidy = 0;
                                //PreModels.Add(model);
                                this.Dispatcher.Invoke(new Action(() => { ViewModel.Models.Add( model); }));
                            }
                        }
                        catch (Exception e)
                        {
                            //PreModels.Add(model);
                            this.Dispatcher.Invoke(new Action(() => { ViewModel.Models.Add( model); }));
                        }
                    }
                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        progressBar.Value++;
                    }));
                }
            }));
            task.ContinueWith(t =>
            {
                 this.Dispatcher.BeginInvoke(new Action(() => {
                     //ViewModel.Models = PreModels;
                     if (ViewModel.RootPath != path)
                     {
                         Model_FileSystem model = new Model_FileSystem() { Name = "..", FullPath = path.Substring(0, path.LastIndexOf('\\')) };
                         ViewModel.Models.Insert(0, model);
                     }
                     progressBar.Visibility = Visibility.Collapsed;
                     progressBar.Value = 0;
                 }));
            });
            task.Start();
        }

        /// <summary>
        /// 查看整理失败的文件
        /// </summary>
        //public void RenderView_Fail(string path)
        //{
        //    DirectoryInfo rootDir = new DirectoryInfo(path);
        //    ViewModel.CurPath = path;
        //    ViewModel.Models.Clear();

        //    foreach (FileSystemInfo fsi in rootDir.GetFileSystemInfos())
        //    {
        //        if (fsi.Name.StartsWith("~$"))
        //            continue;
        //        Model_FileSystem model = new Model_FileSystem() { Name = fsi.Name, FullPath = fsi.FullName };
        //        if (fsi.Attributes == FileAttributes.Directory)
        //        {
        //            model.Type = SystemType.Dir;
        //            try
        //            {
        //                int count = Convert.ToInt32(SQLiteDBHelper.ExecuteScalar("select count(*) from db_File where 路径 like '" + model.FullPath + "%' and (提取!='是' or 可读!='是')", null));
        //                if (count != 0)
        //                {
        //                    model.HasTidy = -2;
        //                    ViewModel.Models.Insert(0, model);
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //            }
        //        }
        //        else
        //        {
        //            model.Type = SystemType.File;
        //            try
        //            {
        //                SQLiteDataReader sqliteReader = SQLiteDBHelper.ExecuteReader("select * from db_File where 路径='" + model.FullPath + "' and (提取!='是' or 可读!='是')", null);
        //                if (sqliteReader.HasRows)
        //                {
        //                    model.HasTidy = -1;
        //                    ViewModel.Models.Add(model);
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //            }
        //        }
        //    }
        //    if (ViewModel.RootPath != path)
        //    {
        //        Model_FileSystem model = new Model_FileSystem() { Name = "..", FullPath = path.Substring(0, path.LastIndexOf('\\')) };
        //        ViewModel.Models.Insert(0, model);
        //    }
        //}

        //public void RenderView_Fail_2(string path)
        //{
        //    DirectoryInfo rootDir = new DirectoryInfo(path);
        //    ViewModel.CurPath = path;
        //    ViewModel.Models.Clear();
        //    foreach (FileSystemInfo fsi in rootDir.GetFileSystemInfos())
        //    {
        //        if (fsi.Name.StartsWith("~$"))
        //            continue;
        //        Model_FileSystem model = new Model_FileSystem() { Name = fsi.Name, FullPath = fsi.FullName};
        //        string code = path.Substring(path.LastIndexOf('\\') + 1);
        //        DataTable dt = SQLiteDBHelper.ExecuteDataTable("select 类型,可读,提取 from db_File where 编号='" + code + "' and 文件名='" + fsi.Name + "'", null);
        //        if (dt.Rows[0]["可读"].ToString() != "是" || dt.Rows[0]["提取"].ToString() != "是")
        //        {
        //            model.HasTidy = -1;
        //        }
        //        else
        //            model.HasTidy = 1;
        //        if (dt.Rows[0]["类型"].ToString() == ".pdf")
        //            model.Type = SystemType.PDF;
        //        else if (dt.Rows[0]["类型"].ToString() == ".doc" || dt.Rows[0]["类型"].ToString() == ".docx" || dt.Rows[0]["类型"].ToString() == ".wps")
        //            model.Type = SystemType.Word;
        //        else
        //            model.Type = SystemType.File;
        //        ViewModel.Models.Add(model);
        //    }
        //}

        /// <summary>
        /// 整理失败的任务渲染
        /// </summary>
        /// <param name="code">整理后的文件路径</param>
        public void RenderView_Fail_3(string Path)
        {
            ViewModel.Models.Clear();
            string code=Path.Substring(Path.LastIndexOf('\\')+1);
            string sql = "select 文件名,路径,可读,提取 from db_File where 编号='"+ code +"'";
            DataTable dt = db.ExecuteDataTable(sql, null);
            foreach (DataRow row in dt.Rows)
            {
                Model_FileSystem model = new Model_FileSystem() { Name = row["文件名"].ToString(), FullPath = row["路径"].ToString() };
                if (row["可读"].ToString() != "是" ||row["提取"].ToString() != "是")
                {
                    model.HasTidy = -1;
                }
                else
                    model.HasTidy = 1;
                string extension = model.Name.Substring(model.Name.LastIndexOf('.'));
                if (extension.ToLower() == ".pdf")
                    model.Type = SystemType.PDF;
                else if (extension.ToLower() == ".doc" || extension.ToLower() == ".docx" || extension.ToLower() == ".wps")
                    model.Type = SystemType.Word;
                else
                    model.Type = SystemType.File;
                ViewModel.Models.Add(model);
            }
        }
        #endregion

        /// <summary>
        /// 双击列表中的文件夹打开该文件夹
        /// </summary>
        public void HandleDoubleClick(Object sender, MouseButtonEventArgs e)
        {
            ListViewItem item = sender as ListViewItem;
            Model_FileSystem model = item.DataContext as Model_FileSystem;
            if (model.Type == SystemType.Dir)
            {
                Render(model.FullPath);
            }
            //else
            //{
            //    if(model.HasTidy==0)
            //    model.Checked = !model.Checked;
            //}
        }

        private bool CheckConfig()
        {
            try
            {
                if(string.IsNullOrEmpty(ConfigHelper.GetValue("PDF_FrontNum"))||string.IsNullOrEmpty(ConfigHelper.GetValue("PDF_BackNum"))||string.IsNullOrEmpty(ConfigHelper.GetValue("Word_FrontNum"))||string.IsNullOrEmpty(ConfigHelper.GetValue("Word_BackNum"))
                    ||string.IsNullOrEmpty(ConfigHelper.GetValue("PDF_Summary_PageNum"))||string.IsNullOrEmpty(ConfigHelper.GetValue("Word_Summary_Num"))
                    || string.IsNullOrEmpty(ConfigHelper.GetValue("Front5DocFetchWord")) || string.IsNullOrEmpty(ConfigHelper.GetValue("Front5RemoveWord")) || string.IsNullOrEmpty(ConfigHelper.GetValue("Back5FetchWord")))
                    return false;
                else
                    return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 开始整理所选文件或文件夹
        /// </summary>
        private void Button_Tidy_Click(object sender, RoutedEventArgs e)
        { 
            Task task=null;
            //tokenSource = new CancellationTokenSource();
            if (!CheckConfig())
            {
                ShowLog("整理规则配置有误,请完善配置文件后重新启动程序", 3);
                return;
            }
            //tHelper = new TidyHelper(db, ViewModel, ShowLog, RemoveModelItem);
            //tHelper.tokenSource = tokenSource;
            try
            {
                
                //每次整理得到不同的pdf文件
                if (b4.Style == (Style)this.FindResource("ButtonStyle3"))
                    task = new Task(() => tHelper.DoError());

                else if (radio_muti.IsChecked.Value)
                {
                    task = new Task(() => tHelper.DoMany(), tokenSource.Token);
                    TextLog.WritwLog("构造批量任务");
                }
                else if (radio_only.IsChecked.Value)
                {
                    task = new Task(() => tHelper.DoAlone(), tokenSource.Token);
                }
                View_Work.ItemContainerStyle = null;

                Btn_Tidy.Header = "停止整理";

                Btn_Tidy.Background = new ImageBrush(new BitmapImage(new Uri(System.AppDomain.CurrentDomain.BaseDirectory + "stop.png", UriKind.RelativeOrAbsolute)));
                TextLog.WritwLog("按钮变换背景");
                Btn_Tidy.MouseLeftButtonUp -= Button_Tidy_Click;
                Btn_Tidy.MouseLeftButtonUp += Button_Click_3;

                ViewModel.CurFail = ViewModel.CurSuc = 0;
                sPanel2.Visibility = Visibility.Visible;
                b1.IsEnabled = b2.IsEnabled = b3.IsEnabled = b4.IsEnabled = false;
                ViewModel.CheckBoxEnable = false;
            }
            catch (Exception cep)
            {
                TextLog.WritwLog(cep.Message);
                MessageBox.Show(cep.Message);
            }
            try
            {
                
                task.ContinueWith(t =>
                {
                    Style s = FindResource("itemstyle") as Style;
                    ViewModel.CheckBoxEnable = true;
                    this.Dispatcher.Invoke(
                    new System.Windows.Forms.MethodInvoker(() =>
                    {
                        Btn_Tidy.IsEnabled = true;
                        View_Work.ItemContainerStyle = s;
                    }));
                    
                    this.Dispatcher.Invoke(new Action(() =>
                    {
                        b1.IsEnabled = b2.IsEnabled = b3.IsEnabled = b4.IsEnabled = true;
                        cb_All.IsChecked = false;

                        Btn_Tidy.MouseLeftButtonUp -= Button_Click_3;
                        Btn_Tidy.MouseLeftButtonUp += Button_Tidy_Click;
                        Btn_Tidy.Background = new ImageBrush(new BitmapImage(new Uri(System.AppDomain.CurrentDomain.BaseDirectory + "Tidy.png", UriKind.RelativeOrAbsolute)));
                        Btn_Tidy.Header = "开始整理";
                    }));
                    SQLiteConnection connection = new SQLiteConnection(db.connectionString);
                    connection.Close();
                    TextLog.WritwLog("整理任务执行完毕");
                });

            }
            catch (Exception cep1) {
                TextLog.WritwLog(cep1.Message);
                MessageBox.Show(cep1.Message);
            }
            try
            {
                task.Start();
            }
            catch(Exception eee)
            {
                TextLog.WritwLog("整理过程异常:" + eee.Message);
            }
        }
        /// <summary>
        /// 设置提取页数
        /// </summary>
        private void Button_Set_Click(object sender, RoutedEventArgs e)
        {
            SetUp setup;
            if (IS_CUSTOM_DEFINED)
            {
                setup = new SetUp(tHelper.PDF_FRONT_NUM,tHelper.PDF_BACK_NUM,tHelper.WORD_FRONT_NUM,tHelper.WORD_BACK_NUM);
               
            }
            else
            {
                setup = new SetUp(false);
                
            }
            setup.CancalDifine = CancelCostomDefine;
            setup.ConfirmDifine = CustomDefinePage;
            setup.ShowDialog();
        }

        /// <summary>
        /// 点击全选按钮
        /// </summary>
        private void AllCheckBox_Click(object sender, RoutedEventArgs e)
        {
           
            CheckBox cb = sender as CheckBox;
            if (ViewModel == null)
                return;
            if (cb.IsChecked.Value)
            {
                if (View_Work.SelectedItems.Count > 1)
                {
                    foreach (Model_FileSystem m in View_Work.SelectedItems)
                    {
                         m.Checked = true;
                    }
                }
                else
                {

                    foreach (Model_FileSystem model in ViewModel.Models)
                    {
                        model.Checked = true;
                    }
                }
            }
            else
            {
                if (View_Work.SelectedItems.Count > 1)
                {
                    foreach (var model in View_Work.SelectedItems)
                    {
                        ((Model_FileSystem)model).Checked = false;
                    }
                }
                else
                {
                    foreach (Model_FileSystem model in ViewModel.Models)
                    {
                        model.Checked = false;
                    }
                }
            }
        }
        /// <summary>
        /// checkbox点击事件,判断是否整理过
        /// </summary>
        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            if(cb.IsChecked.Value)
            {
               
                if(Directory.Exists(cb.Tag.ToString()))
                {
                    Queue<string> pathQueue = new Queue<string>();
                    pathQueue.Enqueue(cb.Tag.ToString());
                    //开始循环查找文件，直到队列中无任何子目录
                    while (pathQueue.Count > 0)
                    {
                        DirectoryInfo diParent = new DirectoryInfo(pathQueue.Dequeue());
                        foreach (DirectoryInfo diChild in diParent.GetDirectories())
                            pathQueue.Enqueue(diChild.FullName);
                        foreach (FileInfo fi in diParent.GetFiles())
                        {
                            if (!fi.Name.StartsWith("~$"))
                            {
                                if (db.ExecuteDataTable("select * from db_File where 路径='" + fi.FullName + "'", null).Rows.Count > 0)
                                {
                                    ShowLog(cb.Tag.ToString().Substring(cb.Tag.ToString().LastIndexOf('\\')+1)+":该文件夹中已有文件被整理过,不能打包整理",3);
                                    cb.IsChecked = false;
                                    return;
                                }
                            }
                        }
                    }
                }
            }
               
        }

        
        /// <summary>
        /// 计算目录中所有嵌套的文件的个数
        /// </summary>
        /// <param name="sPathName">文件夹路径</param>
        /// <returns>文件个数</returns>
        public int Traverse(string sPathName)
        {
            //创建一个队列用于保存子目录
            int i=0;
            Queue<string> pathQueue = new Queue<string>();
            pathQueue.Enqueue(sPathName);
            //开始循环查找文件，直到队列中无任何子目录
            while (pathQueue.Count > 0)
            {
                DirectoryInfo diParent = new DirectoryInfo(pathQueue.Dequeue());
                foreach (DirectoryInfo diChild in diParent.GetDirectories())
                    pathQueue.Enqueue(diChild.FullName);
                foreach (FileInfo fi in diParent.GetFiles())
                {
                    //if (!fi.Name.StartsWith("~$") )
                        i++;
                }
            }
            return i;
        }
        /// <summary>
        /// 返回一个文件夹中的所有文件类型
        /// </summary>
        /// <param name="sPathName">文件夹路径</param>
        public List<string> GetAllFileTypes(string sPathName)
        {
            //创建一个队列用于保存子目录
            List<string> list = new List<string>();
            Queue<string> pathQueue = new Queue<string>();
            pathQueue.Enqueue(sPathName);
            //开始循环查找文件，直到队列中无任何子目录
            while (pathQueue.Count > 0)
            {
                DirectoryInfo diParent = new DirectoryInfo(pathQueue.Dequeue());
                foreach (DirectoryInfo diChild in diParent.GetDirectories())
                    pathQueue.Enqueue(diChild.FullName);
                foreach (FileInfo fi in diParent.GetFiles())
                {
                    if (!list.Exists(s=>s==fi.Extension))
                        list.Add(fi.Extension);
                }
            }
            return list;
        }

      

         /// <summary>
        /// 显示导入过程记录
        /// </summary>
        /// <param name="strMsg">信息</param>
        /// <param name="grade">1为普通信息 2为重要信息 3为错误信息</param>
        public void ShowLog(string strMsg,int grade=1)
        {
            strMsg = "[" + DateTime.Now.ToLongTimeString() + "]:" + strMsg + "\n\r";
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (grade == 3)
                    p1.Inlines.Add(new Run(strMsg) { Foreground=Brushes.Red});
                else if (grade == 2)
                    p1.Inlines.Add(new Run(strMsg) { Foreground = Brushes.Blue });
                else
                    p1.Inlines.Add(new Run(strMsg) { Foreground = Brushes.Black });
                LogBox.LineDown();
                LogBox.LineDown();
            }));
        }

        /// <summary>
        /// 点击后开始标记
        /// </summary>
        frmTidy frm;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //if (PdfFileName == tHelper.CurPdfName)
                //    return;
                //string pdfpath = ViewModel.Upload_Path + "\\" + PdfFileName;

                if (frm == null || frm.IsDisposed)
                {
                    TextLog.WritwLog("开始初始化frm");
                    frm = new frmTidy(ViewModel.TaskCode, ViewModel.WorkPath, ViewModel.TaskType, ViewModel.RootPath,ViewModel.GongHao, ResetView); 
                }
                TextLog.WritwLog("初始化frm成功");
                frm.Location = new System.Drawing.Point( (int)this.Left, (int)this.Top);
                frm.Show();
                frm.Activate();
            }
            catch (Exception dsf)
            {
                Utility.Log.TextLog.WritwLog("启动标记失败:" + dsf.Message);
            }
           
        }

        private void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        /// <summary>
        /// 刷新获取新任务
        /// </summary>
        private void GetTask_Click(object sender, RoutedEventArgs e)
        {
            GetTask();
        }

       
        /// <summary>
        /// 查看未整理,已整理,全部或有错文件
        /// </summary>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (ViewModel==null||ViewModel.RootPath == null)
                return;
            Style s2 = (Style)this.FindResource("ButtonStyle2");
            b1.Style = b2.Style = b3.Style =b4.Style= s2;
            l1.Foreground = l2.Foreground = l3.Foreground = l4.Foreground = new SolidColorBrush(Color.FromRgb(180, 180, 180));
            Button b = sender as Button;
            Style s3 = (Style)this.FindResource("ButtonStyle3");
            b.Style = s3;

            string name = (sender as Button).Name;
            Button4Checked = false;
            
            switch (name)
            { 
                case "b1"://全部文件
                    RenderView(ViewModel.RootPath);
                    Render = RenderView;
                    l1.Foreground = b1.Foreground;
                    //View_Work.MouseMove += View_Work_MouseMove;
                    break;
                case "b2"://未整理
                    RenderView_UnTidy(ViewModel.RootPath);
                    Render = RenderView_UnTidy;
                    l2.Foreground = b2.Foreground;
                    //View_Work.MouseMove -= View_Work_MouseMove;
                    break;
                case "b3"://已整理
                    RenderView_HasTidy(ViewModel.RootPath);
                    Render = RenderView_HasTidy;
                    l3.Foreground = b3.Foreground;
                    //View_Work.MouseMove -= View_Work_MouseMove;
                    break;
                case "b4"://整理失败
                    string sql="";
                    switch (ViewModel.TaskType)
                    { 
                        case "纯盘":
                            //sql = "select distinct b.编号,b.[提交否] from db_File a join XW_FileOrderinfo b on a.编号=b.[编号] where (a.提取!='是' or a.可读!='是') ";
                            sql = "select 编号,提交否 from XW_FileOrderinfo where 文件名 is null or 文件名=''";
                            break;
                        case"刊盘":
                            //sql = "select distinct b.编号,b.[保存否] from db_File a join db_State b on a.编号=b.[编号] where (a.提取!='是' or a.可读!='是')";
                            sql = "select a.编号, b.保存否 from XW_FileOrderinfo a join  db_State b on a.编号=b.编号 where 文件名 is null or 文件名=''";
                            break;
                    }
                    SQLiteDataReader row = db.ExecuteReader(sql, null);
                    ViewModel.Models.Clear();
                    while (row.Read())
                    {
                        var ro=row[0].ToString();
                        Model_FileSystem model = new Model_FileSystem() { Type = SystemType.Dir, HasTidy = row[1].ToString() == "是" || row[1].ToString() == "不可做" ? 2 : -2, Extension = "文件夹", Name = ro, FullPath = ViewModel.AfterTydyPath + "\\" + ro };
                        ViewModel.Models.Add(model);
                    }
                    Render = RenderView_Fail_3;
                    l4.Foreground = b4.Foreground;
                    Button4Checked = true;
                    break;
            }

        }
        /// <summary>
        /// 右键点击文件弹出删除菜单
        /// </summary>
        private void View_Work_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (View_Work.SelectedItem != null)
            {
                if (!Button4Checked)
                {
                    BtnCheck.Visibility = Visibility.Collapsed;
                    BtnFail.Visibility = Visibility.Visible;
                    //Popup.Height = 120;
                }
                else
                {
                    BtnCheck.Visibility = Visibility.Visible;
                    BtnFail.Visibility = Visibility.Collapsed;
                    //Popup.Height = 150;
                }
                Popup.IsOpen = true;
            }
        }
        /// <summary>
        /// 文件无法整理,直接计入整理失败
        /// </summary>
        public void PutTaskFail(object sender, RoutedEventArgs e)
        {
            Model_FileSystem model= View_Work.SelectedItem as Model_FileSystem;
            string desDirName = ViewModel.AfterTydyPath + "\\" + ViewModel.Begin_Code;
            if (!Directory.Exists(desDirName))
                Directory.CreateDirectory(desDirName);
            #region 文件夹
            if (model.Type == SystemType.Dir)
            {
                DirectoryInfo dir = new DirectoryInfo(model.FullPath);
                var files = dir.GetFiles();
                if (ViewModel.TaskType == "纯盘")
                {
                    foreach (FileInfo f in files)
                    {
                        File.Copy(f.FullName, desDirName + "\\" + f.Name, true);
                        string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}','{4}',{5},\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, f.Name, f.Extension, "否", "否", 0, f.FullName, 0, "null", desDirName + "\\" + f.Name);
                        db.ExecuteNonQuery(sql, null);
                    }
                    string  str = string.Format("insert into XW_FileOrderinfo(编号,路径,提交否) values('{0}','{1}','否')", ViewModel.Begin_Code, desDirName);
                    db.ExecuteNonQuery(str, null);
                }
                else
                {
                    foreach (FileInfo f in files)
                    {
                        File.Copy(f.FullName, desDirName + "\\" + f.Name, true);
                        string str = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",{6},{7},{8},\"{9}\")", ViewModel.Begin_Code, f.Name, "否", "否", 0, f.FullName, "null", 0, "null", desDirName + "\\" + f.Name);
                        db.ExecuteNonQuery(str, null);

                    }
                    string sql = string.Format("insert into XW_FileOrderinfo(编号,删除字样,保密否,提取页数,小样数,路径) values('{0}','{1}','{2}',{3},{4},'{5}')", ViewModel.Begin_Code, "否", "否", 0, 1, desDirName);
                    db.ExecuteNonQuery(sql, null);
                    sql = "insert into db_State values('" + ViewModel.Begin_Code + "','否')";
                    db.ExecuteNonQuery(sql, null);
                }
                model.HasTidy = -1;
                ViewModel.UnTidy-=files.Count();
                ViewModel.TotalFail+=files.Count();
            }
            #endregion
            #region 单个文件
            else
            {
               
                if (ViewModel.TaskType == "纯盘")
                {
                    File.Copy(model.FullPath, desDirName + "\\" + model.Name, true);
                    string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}','{4}',{5},\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, model.Name, model.Extension, "否", "否", 0, model.FullPath, 0, "null", desDirName + "\\" + model.Name);
                    db.ExecuteNonQuery(sql, null);
                    sql=string.Format("insert into XW_FileOrderinfo(编号,路径,提交否) values('{0}','{1}','否')", ViewModel.Begin_Code, desDirName);
                    db.ExecuteNonQuery(sql, null);
                }
                else
                {
                    File.Copy(model.FullPath, desDirName + "\\" + model.Name, true);
                    TextLog.WritwLog("PutTaskFail1");
                    string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",{6},{7},{8},\"{9}\")", ViewModel.Begin_Code, model.Name, "否", "否", 0, model.FullPath, "null", 0, "null", desDirName + "\\" + model.Name);
                    TextLog.WritwLog(sql);
                    TextLog.WritwLog("db:"+db.connectionString);
                    try
                    {
                        db.ExecuteNonQuery(sql, null);
                    }
                    catch (Exception eee)
                    {
                        TextLog.WritwLog("PutTaskFail exception:"+eee.Message);
                    }
                    sql = string.Format("insert into XW_FileOrderinfo(编号,删除字样,保密否,提取页数,小样数,路径) values('{0}','{1}','{2}',{3},{4},'{5}')", ViewModel.Begin_Code, "否", "否", 0, 1, desDirName);
                    db.ExecuteNonQuery(sql, null);
                    TextLog.WritwLog("PutTaskFail2");
                    sql = "insert into db_State values('" + ViewModel.Begin_Code + "','否')";
                    db.ExecuteNonQuery(sql, null);
                }
                model.HasTidy = -1;
                ViewModel.UnTidy--;
                ViewModel.TotalFail++;
            }
            #endregion
            Popup.IsOpen = false;
            if (ViewModel.TaskType == "纯盘")
            {
                ViewModel.Begin_Code = (Convert.ToInt32(ViewModel.Begin_Code) + 1).ToString();
                ConfigHelper.SetValue("BEGIN_CODE", ViewModel.Begin_Code);
            }
            else
            {
                string lastFive = (Convert.ToInt32(ViewModel.Begin_Code.Substring(ViewModel.Begin_Code.Length - 5)) + 1).ToString();
                if (lastFive.Length < 5)
                {
                    string zero = "";
                    for (int i = 1; i <= 5 - lastFive.Length; i++)
                    {
                        zero += "0";
                    }
                    lastFive = zero + lastFive;
                }
                ViewModel.Begin_Code = ViewModel.TaskCode + lastFive;
                ConfigHelper.SetValue("BEGIN_CODE_Kanpan", ViewModel.Begin_Code);

            }
        }

        /// <summary>
        /// 打开文件标注的对话框
        /// </summary>
        public void CheckTask(object sender, RoutedEventArgs e)
        {
            var model = View_Work.SelectedItem as Model_FileSystem;
            if (model.Type == SystemType.Dir)
            {
                tb_Code.Text = model.Name;
            }
            else
            {
                string code = db.ExecuteScalar("select 编号 from db_File where 路径='" + model.FullPath + "'", null).ToString();
                tb_Code.Text = code;
            }
            
           
            Popup.IsOpen = false;
            Popup_Check.IsOpen = true;
        }

        /// <summary>
        /// 整理失败的任务提交
        /// </summary>
        private void ChunPanSubmit(object sender, RoutedEventArgs e)
        {
            Submit_Chunpan(tb_Code.Text);
        }
        /// <summary>
        /// 纯盘篇发布
        /// </summary>
        /// <param name="code">任务编号</param>
        private void Submit_Chunpan(object code)
        {

            string str = "select 顺序 from db_File where 编号='" + code+"'";
            DataTable dt = db.ExecuteDataTable(str, null);
            if (dt.Rows.Count > 1)
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row[0].ToString() == "")
                    {
                        //如果该编号下文件数量大于一并且有文件没有标记顺序则弹出标记小样窗口
                        ShowXiaoYangPop(code.ToString());
                        return;
                    }
                }
            }
            string sql = "";
            switch (ViewModel.TaskType)
            { 
                case"纯盘":
                    sql = "select 提交否 from XW_FileOrderinfo where 编号='" + code + "'";
                    break;
                case"刊盘":
                    sql = "select 保存否 from db_State where 编号='" + code + "'";
                    break;
            }
            if (db.ExecuteScalar(sql, null).ToString() == "是")
            {
                return;
            }

            tb_back.Text = "提交中...";
            switch (ViewModel.TaskType)
            {
                case "纯盘":
                    if (!cb_year.IsChecked.Value)
                    {
                        try
                        {
                            int year = Convert.ToInt32(tb_year.Text);
                            if (year > 2070 || year < 1900)
                            {
                                tb_back.Text = "学位年度设置错误";
                                return;
                            }
                        }
                        catch (Exception e)
                        {
                            tb_back.Text = "学位年度只能是数字";
                            return;
                        }
                    }
                    
                    try
                    {
                        new SubmitHelper(db).ChunPan_Submit(code, GetBiaoZhu(ViewModel.TaskType), TempDic, ViewModel.WorkPath);
                    }
                    catch (Exception fee)
                    {
                        TextLog.WritwLog(fee.Message, true);
                    }
                    
                    break;
                case "刊盘":
                    try
                    {
                        new SubmitHelper(db).Kanpan_Submit(code, GetBiaoZhu(ViewModel.TaskType),TempDic,ViewModel.WorkPath);
                    }
                    catch (Exception fee)
                    {
                        TextLog.WritwLog(fee.Message, true);
                    }
                    TextLog.WritwLog("更新Viewmodel");
                    break;
            }
            UdpReceive();
        }
        /// <summary>
        /// 显示标记小样的popup
        /// </summary>
        /// <param name="code">编号</param>
        private void ShowXiaoYangPop(string code)
        {
            DataTable dt = db.ExecuteDataTable("select 文件名,路径,顺序,整理路径 from db_File where 编号='" + code + "'", null);
            DataGrid_XiaoY.ItemsSource = dt.DefaultView;
            Popup_XiaoY.IsOpen = true;
        }

        /// <summary>
        /// 获取用户标注信息
        /// </summary>
        /// <param name="tasktype">纯盘or刊盘</param>
        /// <returns></returns>
        private Dictionary<string, object> GetBiaoZhu(string tasktype)
        {
            Dictionary<string, object> dic = new Dictionary<string, object>();
            dic.Add("保密", cb_secret.IsChecked.Value ? "是" : "否");
            dic.Add("删除字样", cb_delete.IsChecked.Value ? "是" : "否");
            dic.Add("滞后上网", cb_delete.IsChecked.Value ? "是" : "否");
            if (tasktype == "刊盘")
            {
                DataTable dt_text = db.ExecuteDataTable("select 摘要 from db_File where 编号='" + tb_Code.Text + "'", null);
                string text = "";
                foreach (DataRow row in dt_text.Rows)
                {
                    text += row[0].ToString() + " ";
                }
                text = text.Replace("\"", "");
                dic.Add("摘要", text);
            }
            if (tasktype == "纯盘")
            {
                string ShouQuanFanKui = string.Empty;
                if (rb_yes.IsChecked.Value)
                {
                    dic.Add("备注", null);
                    ShouQuanFanKui = "是";
                }
                else if (rb_no.IsChecked.Value)
                {
                    dic.Add("备注", cb_备注.Text);
                    ShouQuanFanKui = "否";
                }

                else if (rb_不合格.IsChecked.Value)
                {
                    dic.Add("备注", null);
                    ShouQuanFanKui = "不合格";
                }
                dic.Add("版权反馈", ShouQuanFanKui);
                dic.Add("授权", cb_无授权.IsChecked.Value ? "是" : "否");//反义一下, 由于后续岗位的列名是'无授权','无作者签名'
                dic.Add("签名", cb_无签名.IsChecked.Value ? "是" : "否");  //反义一下, 由于后续岗位的列名是'无授权','无作者签名'
                if (cb_year.IsChecked.Value)
                    dic.Add("学位年度", null);
                else
                    dic.Add("学位年度", tb_year.Text);

                if (rb_硕士.IsChecked.Value)
                    dic.Add("级别", "硕士");
                else if (rb_博士.IsChecked.Value)
                    dic.Add("级别", "博士");
                else if (rb_博士后.IsChecked.Value)
                    dic.Add("级别", "博士后");
                else if (rb_待定.IsChecked.Value)
                    dic.Add("级别", null);
            }


            return dic;
        }


        private void CancelCheck(object sender, RoutedEventArgs e)
        {
            if (((Button)sender).Name == "OK")
            {
                Popup_XiaoY.IsOpen = false;
                return;
            }
            Popup_Check.IsOpen = false;
            cb_delete.IsChecked = cb_secret.IsChecked = cb_无授权.IsChecked = cb_无签名.IsChecked = false;
            rb_硕士.IsChecked = rb_yes.IsChecked = true;
            tb_back.Text = "";
            
        }

        
        /// <summary>
        /// 置不可做
        /// </summary>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                FrmNotDo fn = new FrmNotDo();
                var point =Mouse.GetPosition(e.Source as FrameworkElement);
                //fn.Location = new System.Drawing.Point(Convert.ToInt32(point.X), Convert.ToInt32(point.Y));
                //fn.Location = new System.Drawing.Point(100000, 100000);
                fn.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
                
                if (fn.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    tb_back.Text = "不可做提交中...";
                    switch (ViewModel.TaskType)
                    {
                        case "纯盘":
                            new SubmitHelper(db).ChunPan_Submit_UnRead(tb_Code.Text, fn.info, ViewModel.WorkPath, TempDic);
                            UdpReceive();
                            break;
                        case "刊盘":
                            
                            new SubmitHelper(db).KanPan_Submit_UnRead(tb_Code.Text, fn.info,ViewModel.WorkPath,TempDic);
                            UdpReceive();
                            break;
                    }
                }

            }
            catch (Exception f)
            {
                ShowLog(f.Message,3);
                TextLog.WritwLog(f.Message, true);
            }
        }
        /// <summary>
        /// 用来给子线程使用
        /// </summary>
        public void RemoveModelItem(Model_FileSystem m)
        {
            this.Dispatcher.Invoke(new Action(() => {
                //ViewModel.Models.Remove(m);
                m.HasTidy = 1;
                m.Checked = false;
                ViewModel.TotalSuc = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 提取='是' and 可读='是'", null));
                ViewModel.TotalFail = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 提取='否' or 可读='否'", null));
            }));
        }

        /// <summary>
        /// 向注册表中写入加工助手需要的信息
        /// </summary>
        /// <param name="lineID">任务LineID</param>
        /// <param name="postID">任务PostID</param>
        private void InitSetupPath(string lineID, string postID)
        {
            var path = System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory);
            var key = lineID + "_" + postID;
            var errMsg = string.Empty;
            var toolPath = System.IO.Path.Combine(path, "Manual_Import.exe");
            new Register().WriteRegeditKey("Software\\CNKI\\ToolSetup", Register.RegDomain.CurrentUser, key, toolPath, out errMsg);
        }

        #region 取任务
        public void GetTask()
        {
            try
            {
                Register reg = new Register();
                string Ini_Path = reg.ReadRegeditKey("QueuePath", "software\\CNKI\\Assistant\\Work\\", Register.RegDomain.CurrentUser).ToString();
                
                INIManage ini = new INIManage(Ini_Path);
                if (ini.SectionValues("Task") == null)
                {
                    ShowLog("当前没有待做任务", 2);
                    return;
                }
                MatchTask task = null;
                foreach (string strValue in ini.SectionValues("Task"))
                {
                    string strTask = strValue.Substring(strValue.IndexOf('=') + 1);
                    MatchTask mt = strTask.FromJson<MatchTask>();
                    if (mt.TaskStatus == "0")
                    {
                        task = mt;
                        if (task.StartTime == "")
                        {
                            task.StartTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            string json = task.ToJson();
                            string key = strValue.Substring(0, strValue.IndexOf('='));
                            ini.setKeyValue("Task", key, json);
                        }
                        break;
                    }
                }
                if (task == null)
                {
                    ShowLog("当前没有待做任务", 2);
                    return;
                }
                Utility.Log.TextLog.WritwLog("g1");
                InitSetupPath(task.LineID, task.PostID);

                CreateDirectory(task.WorkPath + "\\temp");
                CreateDirectory(task.WorkPath + "\\整理后");
                if (ViewModel == null || task.Code != ViewModel.TaskCode)
                {
                    ViewModel = new ViewModel_Main();
                    //获得当前登陆工号，不同工号的刊盘编号不同
                    Ini_Path = Ini_Path.Remove(Ini_Path.LastIndexOf('\\'));
                    ViewModel.GongHao = Ini_Path.Substring(Ini_Path.LastIndexOf('\\')+1);
                    ViewModel.TaskCode = task.Code;
                    ViewModel.WorkPath = task.WorkPath;
                    ViewModel.tempPath = task.WorkPath + "\\temp";//提取的非正文页临时存储路径
                    ViewModel.Upload_Path = task.UpPath;
                    ViewModel.TaskType = task.ProcMode;
                    ViewModel.AfterTydyPath = task.WorkPath + "\\整理后";

                    if (ViewModel.TaskType == "纯盘")
                    {
                        ViewModel.Begin_Code = ConfigHelper.GetValue("BEGIN_CODE");
                        try { Convert.ToInt32(ViewModel.Begin_Code); }
                        catch {
                            throw new Exception("纯盘任务编号有误,请联系测试人员初始化");
                        }

                        CreateDB_ChunPan(ViewModel.tempPath + "\\" + ViewModel.TaskCode + ".db");
                    }
                    else if (ViewModel.TaskType == "刊盘")
                    {
                        string KanpanCode = ConfigHelper.GetValue("BEGIN_CODE_Kanpan_" + ViewModel.GongHao);
                        ViewModel.Begin_Code = string.IsNullOrEmpty(KanpanCode) ? task.Code + "00001" : ConfigHelper.GetValue("BEGIN_CODE_Kanpan_" + ViewModel.GongHao);
                        TextLog.WritwLog("领取任务编号:" + ViewModel.Begin_Code);
                        CreateDB(ViewModel.tempPath + "\\" + ViewModel.TaskCode + ".db");
                        cb_无授权.Visibility = cb_无签名.Visibility = Group2.Visibility = Group3.Visibility = Group4.Visibility = Visibility.Collapsed;
                        Group1.Width = 300;
                        Group1.HorizontalAlignment = HorizontalAlignment.Center;
                    }
                    Utility.Log.TextLog.WritwLog("g2");


                    DirectoryInfo workDir = new DirectoryInfo(ViewModel.WorkPath);
                    foreach (DirectoryInfo dir in workDir.GetDirectories())
                    {
                        if (dir.Name != "ArticlePublish" && dir.Name != "ArticleUpload" && dir.Name != "temp" && dir.Name != "upload" && dir.Name != "xml" && dir.Name != "整理后")
                        {
                            ViewModel.RootPath = dir.FullName;
                            break;
                        }

                    }
                    Utility.Log.TextLog.WritwLog("g3");
                    if (ViewModel.RootPath == null)
                    {
                        ShowLog("待整理文件没有放在文件夹中", 3);
                        return;
                    }
                    Task tt = new Task(new Action(() =>
                    {
                        TotalUnRar(ViewModel.RootPath);
                        ViewModel.TotalSuc = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 提取='是' and 可读='是'", null));
                        ViewModel.TotalFail = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 提取!='是' or 可读!='是'", null));
                        ViewModel.Total = Traverse(ViewModel.RootPath);
                        ViewModel.UnTidy = ViewModel.Total - ViewModel.TotalSuc - ViewModel.TotalFail;

                        #region 自定义抽图页数相关
                        tokenSource = new CancellationTokenSource();
                        tHelper = new TidyHelper(db, ViewModel, ShowLog, RemoveModelItem);
                        tHelper.tokenSource = tokenSource;
                        IS_CUSTOM_DEFINED = false;
                        #endregion 自定义抽图页数相关

                        this.Dispatcher.Invoke(new Action(() =>
                        {
                            this.DataContext = ViewModel;
                            //SQLiteDBHelper sqlhelper = db;
                            RenderView(ViewModel.RootPath);
                            Style s3 = (Style)this.FindResource("ButtonStyle3");
                            b1.Style = s3;
                            l1.Foreground = b1.Foreground;
                            l2.Foreground = l3.Foreground = l4.Foreground = new SolidColorBrush(Color.FromRgb(180, 180, 180));
                            CheckDelFile();
                            
                            ShowLog("领取任务成功,接收编号:" + ViewModel.TaskCode + ",任务类型:" + ViewModel.TaskType, 2);
                        }));
                    }));
                    tt.Start();
                }

                else
                {
                    ShowLog("现有任务还没有提交!", 3);
                    RenderView(ViewModel.RootPath);
                    return;
                }
            }
            catch (Exception f)
            {
                Utility.Log.TextLog.WritwLog(f.Message);
                ShowLog(f.Message,3);
            }
        }
        #endregion

        #region 构造本地sqlite数据库
        /// <summary>
        /// 创建刊盘任务临时db文件
        /// </summary>
        /// <param name="dbPath"></param>
        /// <returns></returns>
        private void CreateDB(string dbPath)
        {
            db = new SQLiteDBHelper(dbPath);
            //如果不存在改数据库文件，则创建该数据库文件  
            if (!System.IO.File.Exists(dbPath))
            {
                db.CreateDB(dbPath);
                string strsql = "Create Table db_detail(岗位名称 varchar(10))";
                db.ExecuteNonQuery(strsql, null);
                strsql = @"Create Table XW_FileOrderinfo(编号 varchar(20) PRIMARY KEY,接收编号 varchar(20),到岗时间 datetime,授予单位 nvarchar(50),年度 varchar(4),级别 varchar(20),保密否 nchar(1),滞后上网 nchar(1),
            版权反馈否 nvarchar(10),是否授权 nchar(1),是否签名 nchar(1),删除字样 nchar(1),可拆切 nchar(1),精装 nchar(1),备注 nvarchar(1000),可读否 nchar(1),制作说明 nvarchar(1000),学院名称 nvarchar(50),提取页数 int,小样数 int,路径 varchar(200),论文摘要 Text,岗位名称 varchar(10),文件名 varchar(15))";
                db.ExecuteNonQuery(strsql, null);
                strsql = "Create Table db_File(编号 varchar(20), 文件名 varchar(100),提取 nchar(2),可读 nchar(1),起始页 int,路径 varchar(200),摘要 Text,结束页 int,顺序 int,整理路径 varchar(200))";
                db.ExecuteNonQuery(strsql, null);
                strsql = "Create Table db_State(编号 varchar(20),保存否 nvarchar(5))";
                db.ExecuteNonQuery(strsql, null);
            }
        }
        /// <summary>
        /// 创建纯盘任务临时db文件
        /// </summary>
        private void CreateDB_ChunPan(string dbPath)
        {
            db = new SQLiteDBHelper(dbPath);
            //如果不存在改数据库文件，则创建该数据库文件  
            if (!System.IO.File.Exists(dbPath))
            {
                db.CreateDB(dbPath);
                string strsql = @"Create Table XW_FileOrderinfo(编号 varchar(20) PRIMARY KEY,年度 varchar(4),级别 varchar(10),版权反馈否 nvarchar(10),保密否 nchar(1),滞后上网 nchar(1),是否签名 nchar(1),是否授权 nchar(1),删除字样 nchar(1),备注 nvarchar(1000),提取页数 int,小样数 int,路径 varchar(200),提交否 varchar(5),文件名 varchar(15))";
                db.ExecuteNonQuery(strsql, null);
                strsql = "Create Table db_File(编号 varchar(20), 文件名 varchar(100),类型 varchar(8),可读 nchar(1),提取 nchar(1),起始页 int,路径 varchar(200),结束页 int,顺序 int,整理路径 varchar(200))";
                db.ExecuteNonQuery(strsql, null);
            }
        }
        #endregion

        #region 解压文件
        /// <summary>
        /// 解压文件
        /// </summary>
        /// <param name="zipFilePath">压缩文件名称</param>
        /// <param name="unZipDir">解压后的文件夹地址</param>
        private void UnRarFile(string zipFilePath, out string unZipDir)
        {
            string zipfilename = System.IO.Path.GetFileName(zipFilePath);
            string zipfilepath = System.IO.Path.GetDirectoryName(zipFilePath);
            unZipDir = zipfilepath + "\\" + zipfilename.Remove(zipfilename.LastIndexOf("."));
            if (!Directory.Exists(unZipDir))
            {
                Directory.CreateDirectory(unZipDir);
            }
            using (Stream stream = File.OpenRead(zipFilePath))
            {
                var reader = ReaderFactory.Open(stream);
                while (reader.MoveToNextEntry())
                {
                    if (!reader.Entry.IsDirectory)
                    {
                        reader.WriteEntryToDirectory(unZipDir, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);
                    }
                }
            }
        }
        /// <summary>
        /// 解压所有嵌套压缩文件
        /// </summary>
        /// <param name="dirPath">文件夹路径</param>
        public void TotalUnRar(string dirPath)
        {
            //创建一个队列用于保存子目录
            Queue<string> pathQueue = new Queue<string>();
            pathQueue.Enqueue(dirPath);
            //开始循环查找文件，直到队列中无任何子目录
            while (pathQueue.Count > 0)
            {
                string xxxx = pathQueue.Dequeue();
                DirectoryInfo diParent = new DirectoryInfo(xxxx);
                foreach (DirectoryInfo diChild in diParent.GetDirectories())
                {
                    pathQueue.Enqueue(diChild.FullName);
                }
                foreach (FileInfo fi in diParent.GetFiles())
                {
                    if (fi.Extension == ".rar" || fi.Extension == ".zip" || fi.Extension == ".7z")
                    {
                        ShowLog("发现压缩文件,开始解压" + fi.Name, 1);
                        string Path="";
                        try
                        {
                            UnRarFile(fi.FullName, out Path);
                        }
                        catch (Exception e)
                        {
                            //PutUnRarFail(fi.FullName,fi.Name);
                            //ShowLog(fi.Name+"解压失败,可以在整理失败中查看,"+e.Message, 3);
                            ShowLog(fi.Name + "解压失败,文件有可能丢失,请联系研发处理该异常,否则后果自负." + e.Message, 3);
                            TextLog.WritwLog(fi.Name + "解压失败:"+e.Message);
                            continue;
                        }

                        pathQueue.Enqueue(Path);
                        fi.Delete();
                    }
                }
            }
        }
        /// <summary>
        /// 解压失败的文件直接算整理失败
        /// </summary>
        /// <param name="rarFilePath">压缩文件的全路径名称</param>
        /// <param name="fileName">文件名</param>
        public void PutUnRarFail(string rarFilePath,string fileName)
        {
            string desDirName = ViewModel.AfterTydyPath + "\\" + ViewModel.Begin_Code;
            if (!Directory.Exists(desDirName))
                Directory.CreateDirectory(desDirName);
            if (ViewModel.TaskType == "纯盘")
            {
                File.Copy(rarFilePath, desDirName + "\\" + fileName, true);
                string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}','{4}',{5},\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, fileName, fileName.Substring(fileName.LastIndexOf('.')), "否", "否", 0, rarFilePath, 0, "null", desDirName + "\\" + fileName);
                db.ExecuteNonQuery(sql, null);
                sql = string.Format("insert into XW_FileOrderinfo(编号,路径,提交否) values('{0}','{1}','否')", ViewModel.Begin_Code, desDirName);
                db.ExecuteNonQuery(sql, null);
            }
            else
            {
                File.Copy(rarFilePath, desDirName + "\\" + fileName, true);
                string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, fileName, "否", "否", 0, rarFilePath, "null", 0, "null", desDirName + "\\" + fileName);
                db.ExecuteNonQuery(sql, null);
                sql = string.Format("insert into XW_FileOrderinfo(编号,删除字样,保密否,提取页数,小样数,路径) values('{0}','{1}','{2}',{3},{4},'{5}')", ViewModel.Begin_Code, "否", "否", 0, 1, desDirName);
                db.ExecuteNonQuery(sql, null);
                sql = "insert into db_State values('" + ViewModel.Begin_Code + "','否')";
                db.ExecuteNonQuery(sql, null);
            }
            
            if (ViewModel.TaskType == "纯盘")
            {
                ViewModel.Begin_Code = (Convert.ToInt32(ViewModel.Begin_Code) + 1).ToString();
                ConfigHelper.SetValue("BEGIN_CODE", ViewModel.Begin_Code);
            }
            else
            {
                string lastFive = (Convert.ToInt32(ViewModel.Begin_Code.Substring(ViewModel.Begin_Code.Length - 5)) + 1).ToString();
                if (lastFive.Length < 5)
                {
                    string zero = "";
                    for (int i = 1; i <= 5 - lastFive.Length; i++)
                    {
                        zero += "0";
                    }
                    lastFive = zero + lastFive;
                }
                ViewModel.Begin_Code = ViewModel.TaskCode + lastFive;
                ConfigHelper.SetValue("BEGIN_CODE_Kanpan", ViewModel.Begin_Code);

            }
        }
        #endregion

        #region udp消息接收
        public void UdpReceive()
        {

            PublicTool.localUdp.BeginReceive(new AsyncCallback(FinalReceiveCallback), null);
        }
        public void FinalReceiveCallback(IAsyncResult iar)
        {
            if (iar.IsCompleted)
            {
                IPEndPoint remoteEP = PublicTool.GetRemoteEp();
                Byte[] receiveBytes = PublicTool.localUdp.EndReceive(iar, ref remoteEP);
                string FinalBackString = Encoding.GetEncoding("GB2312").GetString(receiveBytes);
                #region 处理过程
                try
                {
                    if (FinalBackString == "Y")
                    {
                        PublicTool.localUdp.BeginReceive(new AsyncCallback(FinalReceiveCallback), null);
                    }
                    else
                    {
                        SubmitResult sr = FinalBackString.FromJson<SubmitResult>();
                        string code = sr.ArticleCode;
                        if (sr.Statu.ToUpper() == "TRUE")
                        {
                            string isRead = TempDic[code];
                            #region 不可做任务
                            //if (sTask.IsRead == "否")
                            //{
                            //    this.Dispatcher.BeginInvoke(new Action(() =>
                            //    {
                            //        string sql = "update XW_FileOrderinfo set 提交否='不可做' where 编号='" + code + "'";
                            //        db.ExecuteNonQuery(sql, null);
                            //        TempDic.Remove(code);
                            //        tb_back.Text = "置不可做成功!";
                            //        if (ViewModel.Models[0].Type == SystemType.Dir)
                            //        {
                            //            ViewModel.Models.Remove(ViewModel.Models.Where(m => m.Name == code).First());
                            //        }
                            //        else
                            //        {
                            //            foreach (var model in ViewModel.Models)
                            //                model.HasTidy = 2;
                            //        }
                            //    }));
                            //}
                            #endregion
                            #region 正常任务
                            //else
                            //{
                            //    this.Dispatcher.BeginInvoke(new Action(() =>
                            //    {
                            //        string sql = string.Format("update XW_FileOrderinfo set 年度='{0}',级别='{1}',保密否='{2}',版权反馈否='{3}',是否签名='{4}',是否授权='{5}',备注='{6}',删除字样='{7}',提交否='是' where 编号='{8}'",
                            //           sTask.Year, sTask.Level, sTask.IsSecret, sTask.Iscopyright, sTask.IsQM, sTask.IsSQ, sTask.Explain, sTask.DeleteWords, code);
                            //        db.ExecuteNonQuery(sql, null);
                            //        TempDic.Remove(code);
                            //        tb_back.Text = "提交成功!";
                            //        if (ViewModel.Models[0].Type == SystemType.Dir)
                            //        {
                            //            ViewModel.Models.Remove(ViewModel.Models.Where(m => m.Name == code).First());
                            //        }
                            //        else
                            //        {
                            //            foreach (var model in ViewModel.Models)
                            //                model.HasTidy = 2;
                            //        }
                            //    }));
                            //}
                            #endregion

                            #region 新的提交方法
                            this.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                if (ViewModel.TaskType == "纯盘")
                                {
                                    switch (isRead)
                                    {
                                        case "是":
                                            db.ExecuteNonQuery("update XW_FileOrderinfo set 提交否='是' where 编号='" + code + "'", null);
                                            tb_back.Text = "提交成功!";
                                            break;
                                        case "否":
                                            db.ExecuteNonQuery("update XW_FileOrderinfo set 提交否='不可做' where 编号='" + code + "'", null);
                                            tb_back.Text = "置不可做成功!";
                                            break;
                                    }
                                }
                                else if (ViewModel.TaskType == "刊盘")
                                {
                                    switch (isRead)
                                    {
                                        case "是":
                                            db.ExecuteNonQuery("update db_State set 保存否='是' where 编号='" + code + "'", null);
                                            tb_back.Text = "提交成功!";
                                            break;
                                        case "否":
                                            db.ExecuteNonQuery("update db_State set 保存否='不可做' where 编号='" + code + "'", null);
                                            db.ExecuteNonQuery("update XW_FileOrderinfo set 可读否='否' where 编号='" + code + "'", null);
                                            tb_back.Text = "置不可做成功!";
                                            break;
                                    }
                                }
                                
                                TempDic.Remove(code);
                                
                                if (ViewModel.Models[0].Type == SystemType.Dir)
                                {
                                    ViewModel.Models.Remove(ViewModel.Models.Where(m => m.Name == code).First());
                                }
                                else
                                {
                                    foreach (var model in ViewModel.Models)
                                        model.HasTidy = 2;
                                }
                            }));
                            #endregion
                        }
                        else
                        {
                            this.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                tb_back.Text = sr.ErrInfo;
                                TextLog.WritwLog(code + "篇发布失败:" + sr.ErrInfo);
                            }));
                        }
                    }
                }
                catch (Exception e)
                {
                    TextLog.WritwLog(e.Message);
                    string code = FinalBackString.Substring(FinalBackString.IndexOf(':') + 1, FinalBackString.IndexOf(',') - FinalBackString.IndexOf(':') - 1);
                    code = code.Remove(0, 1);
                    code = code.Substring(0, code.Length - 1);
                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        //listView2.Items[code].SubItems[1].Text = "提交失败";
                    }));
                }
                #endregion
                //receiveDone.Set();

            }

        }
        #endregion

        #region 删除、打开文件
        /// <summary>
        /// 删除文件
        /// </summary>
        private void DeleteFile(object sender, RoutedEventArgs e)
        {
            var models = ViewModel.Models.Where(m => m.Checked);
            if (models.Where(m => m.Type == SystemType.Dir).Count() > 0)
            {
                MessageBox.Show("为防止任务异常,只能删除文件", "禁止删除文件夹");
                return;
            }
            else
            {
                var model = View_Work.SelectedItem as Model_FileSystem;
                if (model.Type == SystemType.Dir)
                { 
                    MessageBox.Show("为防止任务异常,只能删除文件", "禁止删除文件夹");
                    return;
                }

            }
            MessageBoxResult mbr= MessageBox.Show("删除文件后无法找回,确定删除吗?", "警告", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (mbr == MessageBoxResult.OK)
            {
                
                if (models.Count() > 0)
                {
                    //foreach (Model_FileSystem mo in models)
                    for(int i=0;i<models.Count();i++)
                    {
                        //if (models.ElementAt(i).Type == SystemType.Dir)
                        //    Directory.Delete(models.ElementAt(i).FullPath, true);
                        //else
                            File.Delete(models.ElementAt(i).FullPath);
                            Db_Delte_Sync(models.ElementAt(i).FullPath, FileAttributes.Normal);
                        ViewModel.Models.Remove(models.ElementAt(i));
                    }
                    

                }
                else
                {
                    var model = View_Work.SelectedItem as Model_FileSystem;
                    //if (model.Type == SystemType.Dir)
                    //    Directory.Delete(model.FullPath, true);
                    //else
                        File.Delete(model.FullPath);
                        Db_Delte_Sync(model.FullPath, FileAttributes.Normal);
                    ViewModel.Models.Remove(model);
                }
                ViewModel.TotalSuc = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 提取!='否' and 可读!='否'", null));
                ViewModel.TotalFail = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 提取='否' or 可读='否'", null));
                ViewModel.Total = Traverse(ViewModel.RootPath);
                ViewModel.UnTidy = ViewModel.Total - ViewModel.TotalSuc - ViewModel.TotalFail;
                Popup.IsOpen = false;
            }
        }

        private void Db_Delte_Sync(string fullpath,FileAttributes attr)
        {
            if (attr == FileAttributes.Normal)
            {
                DataTable dt= db.ExecuteDataTable("select 编号 from db_File where 编号=(select 编号 from db_File where 路径='" + fullpath + "')", null);
                if (dt.Rows.Count == 1)//如果该任务只包含这一个文件,则要把对应任务也删除
                {
                    string code = dt.Rows[0][0].ToString();
                    db.ExecuteNonQuery("delete from XW_FileOrderinfo where 编号='"+code+"'", null);
                    db.ExecuteNonQuery("delete from db_File where 路径 like '" + fullpath + "%'", null);
                    if(ViewModel.TaskType=="刊盘")
                        db.ExecuteNonQuery("delete from db_State where 编号='" + code + "'", null);
                }
                else if (dt.Rows.Count > 1)
                {
                    db.ExecuteNonQuery("delete from db_File where 路径 like '" + fullpath + "%'", null);
                }

            }
            else if (attr == FileAttributes.Directory)
            { }
        }
        /// <summary>
        /// 按文件类型批量删除文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DeleteFileByType(object sender, RoutedEventArgs e)
        {
            if (!CheckDelFile())
            {
                var types = GetAllFileTypes(ViewModel.RootPath);
                DelPanel.Children.Clear();
                foreach (var t in types)
                {
                    CheckBox cb = new CheckBox();
                    cb.Content = t;
                    cb.Margin = new Thickness(5, 5, 0, 5);
                    DelPanel.Children.Add(cb);
                }
                DeletePop.IsOpen = true;
            }
            Popup.IsOpen = false;
        }

        private void CancelDelete(object sender, RoutedEventArgs e)
        {
            DeletePop.IsOpen = false;
        }
        /// <summary>
        /// 批量删除文件
        /// </summary>
        private void ConfirmDelete(object sender, RoutedEventArgs e)
        {
            List<string> types = new List<string>();
            foreach (var ele in DelPanel.Children)
            {
                CheckBox cb = ele as CheckBox;
                if (cb.IsChecked.Value)
                    types.Add(cb.Content.ToString());
            }

            var models = ViewModel.Models.Where(m => types.Exists(t =>
            {
                if (m.Name.LastIndexOf('.') > 0)
                    return t == m.Name.Substring(m.Name.LastIndexOf('.'));
                else
                    return false;
            })).ToList();
            for (int i = 0; i < models.Count(); i++)
            {
                ViewModel.Models.Remove(models[i]);
            }


            Queue<string> pathQueue = new Queue<string>();
            pathQueue.Enqueue(ViewModel.RootPath);
            //开始循环查找文件，直到队列中无任何子目录
            while (pathQueue.Count > 0)
            {
                DirectoryInfo diParent = new DirectoryInfo(pathQueue.Dequeue());
                foreach (DirectoryInfo diChild in diParent.GetDirectories())
                    pathQueue.Enqueue(diChild.FullName);
                foreach (FileInfo fi in diParent.GetFiles())
                {
                    if (types.Exists(t => t == fi.Extension))
                    {
                        fi.Delete();
                    }
                }
            }

            ShowLog("删除文件成功!");
            ViewModel.TotalSuc = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 提取='是' and 可读='是'", null));
            ViewModel.TotalFail = Convert.ToInt32(db.ExecuteScalar("select count(*) from db_File where 提取='否' or 可读='否'", null));
            ViewModel.Total = Traverse(ViewModel.RootPath);
            ViewModel.UnTidy = ViewModel.Total - ViewModel.TotalSuc - ViewModel.TotalFail;
            DeletePop.IsOpen = false;
        }
        /// <summary>
        /// 打开文件
        /// </summary>
        public void OpenFile(object sender, RoutedEventArgs e)
        {
            var model = View_Work.SelectedItem as Model_FileSystem;
            System.Diagnostics.Process.Start(model.FullPath);
        }
        #endregion

        #region 为鼠标框选添加相应事件
        //private void Window_Loaded(object sender, RoutedEventArgs e)
        //{
        //   AdornerLayer.GetAdornerLayer(View_Work).Add(myAdorner);
        //}

        private void View_Work_MouseMove(object sender, MouseEventArgs e)
        {
            if (myDragStartPoint.HasValue)
            {
                Rect r = new Rect(myDragStartPoint.Value, e.GetPosition(View_Work) - myDragStartPoint.Value);
                //myAdorner.HighlightArea = r;
                List<Model_FileSystem> items = View_Work.GetItemAt<Model_FileSystem>(r);
                //if (items.Where(m => m.Checked == false).Count() == 0)
                //{
                //    foreach (var i in items)
                //    {
                //        View_Work.SelectedItems.Add(i);
                //        i.Checked = false;
                //    }
                //}
                if (items.Count > 0)
                {
                    View_Work.SelectedItems.Clear();
                    foreach (var i in items)
                    {
                        View_Work.SelectedItems.Add(i);
                    }
                }
                //else
                //    View_Work.SelectedItems.Clear();
            }
        }

        private void View_Work_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                myDragStartPoint = e.GetPosition(View_Work);
            }
        }

        private void View_Work_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                myDragStartPoint = null;

                //myAdorner.HighlightArea = new Rect();
            }
        }

        //private void View_Work_MouseLeave(object sender, MouseEventArgs e)
        //{
        //    myAdorner.HighlightArea = new Rect();
        //}
        #endregion

        #region 鼠标移动popup
        [DllImport("user32")]
        public static extern IntPtr WindowFromPoint(POINT Point);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImportAttribute("user32")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImportAttribute("user32")]
        public static extern bool ReleaseCapture();

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        private void lblCaption_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            POINT curPos;
            IntPtr hWndPopup;

            GetCursorPos(out curPos);
            hWndPopup = WindowFromPoint(curPos);

            ReleaseCapture();
            SendMessage(hWndPopup, WM_NCLBUTTONDOWN, new IntPtr(HT_CAPTION), IntPtr.Zero);
        }
        #endregion 鼠标移动popup

        #region 窗体基本操作
        //最小化
        private void btnMin_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = System.Windows.WindowState.Minimized;
        }
        //最大化
        private void btnMax_Click(object sender, RoutedEventArgs e)
        {
            if (this.WindowState != System.Windows.WindowState.Maximized)
            {
                this.WindowState = System.Windows.WindowState.Maximized;
                ColumnName.Width = 500;
            }
            else
            {
                this.WindowState = System.Windows.WindowState.Normal;
                ColumnName.Width = 200;
            }
        }
        //关闭
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public void DragWindow(object sender, MouseButtonEventArgs args)
        {
            this.DragMove();
        }
        #endregion

        #region 标记窗口基本操作
        private void rb_不合格_Checked(object sender, RoutedEventArgs e)
        {
                cb_无签名.IsChecked = false;
                cb_无签名.IsEnabled = false;
        }
        private void rb_不合格_Unchecked(object sender, RoutedEventArgs e)
        {
            cb_无签名.IsEnabled = true;
        }
        private void rb_no_Checked(object sender, RoutedEventArgs e)
        {
            cb_备注.IsEnabled = true;
        }
        private void rb_no_Unchecked(object sender, RoutedEventArgs e)
        {
            cb_备注.IsEnabled = false;
        }
        private void cb_无授权_Click(object sender, RoutedEventArgs e)
        {
            if (cb_无授权.IsChecked.Value)
            {
                rb_yes.IsChecked = true;
                Group4.IsEnabled = false;
                cb_无签名.IsEnabled = false;
            }
            else
            {
                Group4.IsEnabled = true;
                cb_无签名.IsEnabled = true;
            }
        }
        private void cb_year_Click(object sender, RoutedEventArgs e)
        {
            if (cb_year.IsChecked.Value)
            {
                tb_year.IsEnabled = false;
            }
            else
            {
                tb_year.IsEnabled = true;
            }
        }
        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            tb_keyWord.Opacity = 1;
            btn_find.Opacity = 0.8;
        }
        private void tb_keyWord_LostFocus(object sender, RoutedEventArgs e)
        {
            tb_keyWord.Opacity = 0.1;
            btn_find.Opacity = 0;
        }
        private void btn_find_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel == null || ViewModel.Models == null)
                return;
            ViewModel.Models.Clear();
            Queue<string> pathQueue = new Queue<string>();
            pathQueue.Enqueue(ViewModel.RootPath);
            //开始循环查找文件，直到队列中无任何子目录
            string keyword=tb_keyWord.Text;
            while (pathQueue.Count > 0)
            {
                DirectoryInfo diParent = new DirectoryInfo(pathQueue.Dequeue());
                foreach (DirectoryInfo diChild in diParent.GetDirectories())
                    pathQueue.Enqueue(diChild.FullName);
                foreach (FileInfo fsi in diParent.GetFiles())
                {
                    if (fsi.Name.Contains(keyword))
                    {
                        Model_FileSystem model = new Model_FileSystem() { Name = fsi.Name, FullPath = fsi.FullName, Time = fsi.LastWriteTime.ToString("yyyy/MM/dd hh:mm") };
                        model.FileSize = ((FileInfo)fsi).Length / 1024 + "KB";
                        model.HasTidy = GetFileState(model.FullPath);
                        if (fsi.Extension.ToLower() == ".doc" || fsi.Extension.ToLower() == ".docx" || fsi.Extension.ToLower() == ".wps")
                            model.Type = SystemType.Word;
                        else if (fsi.Extension.ToLower() == ".pdf")
                            model.Type = SystemType.PDF;
                        else
                            model.Type = SystemType.File;
                        model.Extension = fsi.Extension.Substring(1).ToUpper() + "文件";
                        ViewModel.Models.Add(model);
                    }
                }
            }
        }
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            tokenSource.Cancel();
            //Btn_Tidy.MouseLeftButtonUp -= Button_Click_3; 
            //Btn_Tidy.MouseLeftButtonUp += Button_Tidy_Click;
            //Btn_Tidy.Background = new ImageBrush(new BitmapImage(new Uri("Tidy.png", UriKind.Relative)));
            //Btn_Tidy.Header = "开始整理";
        }
        #endregion
        /// <summary>
        /// 标记小样
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGrid_XiaoY_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                var tb=e.EditingElement as TextBox;
                try
                {
                    var Rv = e.Row.Item as DataRowView;
                    string path = Rv.Row["路径"].ToString();
                    string Cur_TaskCode = tb_Code.Text;
                    string sql = ViewModel.TaskType == "纯盘" ? "select 提交否 from XW_FileOrderinfo where 编号=(select 编号 from  db_File  where 路径='" + path + "')" :
                        "select 保存否 from db_State where 编号=(select 编号 from  db_File  where 路径='" + path + "')";
                    if (db.ExecuteScalar(sql, null).ToString() == "是")
                    {
                        MessageBox.Show("文件已上传提交,无法标注顺序");
                        tb.Text = db.ExecuteScalar("select 顺序 from  db_File  where 路径='" + path + "'", null).ToString();
                        return;
                    }
                    string newName = tb.Text;
                    try
                    {
                        int num = Convert.ToInt32(newName);
                        if (num < 1 || num > 99)
                        {
                            MessageBox.Show("小样编号必须是小于100的正数", "Error");
                            tb.Text = db.ExecuteScalar("select 顺序 from  db_File  where 路径='" + path + "'", null).ToString();
                            return;
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("小样编号必须是数字", "Error");
                        tb.Text = db.ExecuteScalar("select 顺序 from  db_File  where 路径='" + path + "'", null).ToString();
                        return;
                    }
                    string afterPath = Rv.Row["整理路径"].ToString();
                    string newPath;
                    if (File.Exists(afterPath))
                    {
                        FileInfo oldFile = new FileInfo(afterPath);
                        newPath = oldFile.Directory.FullName + "\\" + newName + "_" + oldFile.Name;
                        oldFile.MoveTo(newPath);
                    }
                    else
                    {
                        string Cur_filename = afterPath.Substring(afterPath.LastIndexOf('\\') + 1);
                        FileInfo file = new FileInfo(ViewModel.WorkPath + "\\ArticleUpload\\" + Cur_TaskCode + "\\" + Cur_filename);
                        newPath = file.Directory.FullName + "\\" + newName + "_" + file.Name;
                        file.MoveTo(newPath);
                    }
                    db.ExecuteNonQuery("update db_File set 整理路径='" + newPath + "',顺序='" + newName + "' where 路径='" + path + "'", null);
                    Rv.Row["整理路径"] = newPath;
                }
                catch (InvalidExpressionException sq)
                {
                    MessageBox.Show(sq.Message);
                }
            }
            
        }

        /// <summary>
        /// 按文件类型归类,提示用户删除文件
        /// </summary>
        private bool CheckDelFile()
        {
            var dic_type = GetAllDelFile(ViewModel.RootPath);
            if (dic_type.Count == 0)
                return false;
            FileTypeDel Win = new FileTypeDel(db,ViewModel.TaskType);
            foreach (string s in dic_type.Keys)
            {
                TabItem ti = new TabItem();
                ti.Header = s;
                DataGrid grid = new DataGrid();
                grid.AutoGenerateColumns = false;
                grid.Columns.Add(new DataGridTextColumn() { Header = "文件名" ,Binding=new Binding("Name"),MinWidth=200});
                grid.Columns.Add(new DataGridTextColumn() { Header = "文件路径", Binding = new Binding("FullPath")});
                grid.ItemsSource = dic_type[s];

                ti.Content = grid;
                Win.Tab_File.Items.Add(ti);
            }
            Win.Top = this.Top;
            Win.Left = this.Left;
            Win.Show();
            return true;
        }

        /// <summary>
        /// 文件夹路径
        /// </summary>
        /// <param name="sPathName"></param>
        /// <returns>后缀名为key,文件集合为value的字典</returns>
        public Dictionary<string, List<DelFile>> GetAllDelFile(object sPathName)
        {
            //创建一个队列用于保存子目录
            string[] StandadTypes = { ".doc", ".docx", ".pdf" };
            Dictionary<string, List<DelFile>> dic = new Dictionary<string, List<DelFile>>();
            Queue<string> pathQueue = new Queue<string>();
            pathQueue.Enqueue(sPathName.ToString());
            //开始循环查找文件，直到队列中无任何子目录
            while (pathQueue.Count > 0)
            {
                DirectoryInfo diParent = new DirectoryInfo(pathQueue.Dequeue());
                foreach (DirectoryInfo diChild in diParent.GetDirectories())
                    pathQueue.Enqueue(diChild.FullName);
                List<DelFile> list = new List<DelFile>();
                foreach (FileInfo fi in diParent.GetFiles())
                {
                    string ext = fi.Extension.ToLower();
                    if (StandadTypes.Contains(ext))
                        continue;
                    if (dic.ContainsKey(ext))
                    {
                        dic[ext].Add(new DelFile(fi.Name, fi.DirectoryName));
                    }
                    else
                    {
                        List<DelFile> list_file = new List<DelFile>();
                        list_file.Add(new DelFile(fi.Name, fi.DirectoryName));
                        dic.Add(ext, list_file);
                    }
                }
            }
            return dic;
        }
        /// <summary>
        /// 自定义抽图参数
        /// </summary>
        public void CustomDefinePage(int pdffront, int pdfback, int wordfront, int wordback)
        {
            if (tHelper == null)
            {
                ShowLog("自定义抽图参数失败,请先领取任务",3);
            }
            else
            {
                tHelper.PDF_FRONT_NUM = pdffront;
                tHelper.PDF_BACK_NUM = pdfback;
                tHelper.WORD_FRONT_NUM = wordfront;
                tHelper.WORD_BACK_NUM = wordback;
                tHelper.CUSTOM_DEFINE = true;
                IS_CUSTOM_DEFINED = true;
                string format="自定义抽图参数,提取pdf前{0}页,后{1}页,提取word前{2}页,后{3}页";
                ShowLog(string.Format(format, pdffront, pdfback, wordfront, wordback), 2);
            }
        }

        public void CancelCostomDefine()
        {
            tHelper.PDF_FRONT_NUM = Convert.ToInt32(ConfigHelper.GetValue("PDF_FrontNum"));
            tHelper.PDF_BACK_NUM = Convert.ToInt32(ConfigHelper.GetValue("PDF_BackNum"));
            tHelper.WORD_FRONT_NUM = Convert.ToInt32(ConfigHelper.GetValue("Word_FrontNum"));
            tHelper.WORD_BACK_NUM = Convert.ToInt32(ConfigHelper.GetValue("Word_BackNum"));
            tHelper.CUSTOM_DEFINE = false;
            IS_CUSTOM_DEFINED = false;
            ShowLog("抽图参数恢复默认值", 2);
        }
    }
}