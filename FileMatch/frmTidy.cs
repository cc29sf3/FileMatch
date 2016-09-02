using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using Word=Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office;
using DocumentFormat.OpenXml.Office2010;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using System.Xml.Linq;
using Newtonsoft.Json;
using FileMatch.Helper;
using Utility.Common;
using FileMatch.Entity;
using Utility.Dao;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using Utility.Log;
using Utility.Submit;
using System.Runtime.InteropServices;

namespace FileMatch
{
    public partial class frmTidy : Form
    {
        string Work_Path="";//工作路径
        string Upload_Path = "";//上传路径
        string unit = "";//授予单位
        string Root_Path = "";//待整理文件的根目录
        string TaskType = "";
        string GongHao = "";
        Action<object> PianSubmit;//篇提交方法
        Action ReloadTask;//重载下一任务
        Thread backgroudThread;//后台线程加载任务列表
        const int WM_KEYDOWN = 256;
        const int WM_SYSKEYDOWN = 260;
      
        //string Cur_TaskCode;//当前任务编号
        int totalPageCount;//pdf总页数
        int totalTaskCount=0;//总任务数

        SubmitHelper sHelper;
        SQLiteDBHelper dbHelper;
        //切换下一任务的pdf文件,code是要切换的任务编号
        public delegate void PdfFileChangeEventHandler(string code);
        //提交前验证事件,code是要验证的任务编号
        public delegate bool SubmitCheckHandler(string code);
        public event PdfFileChangeEventHandler PdfFileChangeEvent;
        public event SubmitCheckHandler SubmitCheckEvent;

        public IPEndPoint remoteEP = null;
        Dictionary<string, string> TempDic;//查看某个任务是否可做

        [DllImport("user32")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint control, Keys vk);
        //注册热键的api    
        [DllImport("user32")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private void Form1_Load(object sender, EventArgs e)
        {
            //注册热键(窗体句柄,热键ID,辅助键,实键)   
            RegisterHotKey(this.Handle, 888, 0, Keys.Down);
            RegisterHotKey(this.Handle, 777, 0, Keys.Up);
            RegisterHotKey(this.Handle, 666, 0, Keys.PageDown);
            RegisterHotKey(this.Handle, 555, 0, Keys.Delete);
        }
      
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case 0x0312:    //这个是window消息定义的   注册的热键消息    
                    if (m.WParam.ToString().Equals("888"))
                        axCAJAX1.TurnToPage(1, 2);
                    else if (m.WParam.ToString().Equals("777"))
                        axCAJAX1.TurnToPage(1, 1);
                    else if (m.WParam.ToString().Equals("666"))
                        ChangeNextPDF();
                    else if (m.WParam.ToString().Equals("555"))
                        DeleteXiaoYang();
                    break;
            }
            base.WndProc(ref m);
        }  

        public frmTidy()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="code">接收编号</param>
        /// <param name="pathwork">工作路径</param>
        /// <param name="taskType">任务类型</param>
        /// <param name="updateMethod">更新主窗体文件状态的方法</param>
        /// <param name="getCurUsePdf">获取当前正在整理使用的非正文页pdf文件名</param>
        /// <param name="rootPath">待整理文件根路径</param>
        /// <param name="gongHao">工号</param>
        /// <param name="resetView">设置完成标记后重载任务</param>
        public frmTidy(string code, string pathwork, string taskType, string rootPath,string gongHao,Action resetView)
        {
           
            InitializeComponent();

            TextLog.WritwLog("开始启动endpoint");
            remoteEP = PublicTool.GetRemoteEp();
            TextLog.WritwLog("启动endpoint成功");
            Work_Path = pathwork;
            Root_Path = rootPath;
            ReloadTask = resetView;
            
            string filename = Work_Path + "\\Explain.xml";
            unit = XDocument.Load(filename).Element("ExplainInfo").Element("Info").Value.Replace("授予单位:", "");
            tb_School.Text = unit;
            tb_TaskCode.Text = code;
            Upload_Path = Work_Path+@"\upload";
            TaskType = taskType;
            GongHao = gongHao;
            PdfFileChangeEvent += UpdateUi;
            dbHelper = new SQLiteDBHelper(pathwork + "\\temp\\" + code + ".db");
            sHelper = new SubmitHelper(dbHelper);
            if (taskType == "刊盘")
            {
                PianSubmit = SaveInfo;
                dataGridView1.CellEndEdit += dataGridView1_CellEndEdit_Chunpan;
                SubmitCheckEvent += KanPanCheck;
                TempDic = new Dictionary<string, string>();
            }
            else if (taskType == "纯盘")
            {
                PianSubmit = Submit_Chunpan;
                dataGridView1.CellEndEdit += dataGridView1_CellEndEdit_Chunpan;
                SubmitCheckEvent += ChunPanCheck;
                TempDic = new Dictionary<string, string>();
            }
            
            try
            {
                Initial();
                loadListView();
                UdpReceive();
                backgroudThread = new Thread(new ThreadStart(BackThreadStart));
                backgroudThread.IsBackground = true;
                backgroudThread.Start();
                GetCurTask();
                
            }
            catch (Exception e)
            {
                TextLog.WritwLog(e.Message);
            }
          
        }

        /// <summary>
        /// 后台监视是否有整理好的任务,加载到标记页面
        /// </summary>
        public void BackThreadStart()
        {
            while (true)
            {
                try
                {
                    Thread.Sleep(10000);
                    this.Invoke(new Action(() =>
                    {
                        string sql = "";
                        if (listView2.Items.Count == 0)
                        {
                            sql = "select 编号 from XW_FileOrderinfo where 文件名 is not null and 文件名 !=''";
                        }
                        else
                        {
                            string lastCode = listView2.Items[listView2.Items.Count - 1].Text;
                            sql = "select 编号 from XW_FileOrderinfo where 编号>'" + lastCode + "' and 文件名 is not null and 文件名 !=''";
                        }
                        DataTable dt = dbHelper.ExecuteDataTable(sql, null);
                        foreach (DataRow r in dt.Rows)
                        {
                            ListViewItem lvi = new ListViewItem(r[0].ToString()) { Name = r[0].ToString(), ImageIndex = 1 };
                            lvi.SubItems.Add(new ListViewItem.ListViewSubItem());
                            listView2.Items.Add(lvi);
                        }
                    }));
                    
                }
                catch (Exception e)
                {
                    TextLog.WritwLog("后台线程运行错误:"+e.Message);
                }
            }
        }
      
        /// <summary>
        /// 添加文件类型筛选事件
        /// </summary>
        private void Initial()
        {
            if (TaskType == "纯盘")
            {
                button1.Click += Total_Submit_ChunPan;
            }
            else
            {
                button1.Click += Total_Submit_KanPan;
                groupBox5.Visible=groupBox_Grade.Visible = groupBox_Year.Visible = cb_authoration.Visible = cb_signature.Visible = false;
            }

        }
        /// <summary>
        /// 获取第一个待处理篇任务,同步ui界面
        /// </summary>
        private void GetCurTask()
        {
            string Cur_TaskCode="";
            if (TaskType == "刊盘")
            {
                Cur_TaskCode = dbHelper.ExecuteScalar("select a.编号 from db_State a join XW_FileOrderinfo b on a.编号=b.编号  where a.保存否='否' and b.文件名!='' and not exists(select * from db_File where 编号=a.[编号] and (可读='否' or 提取='否'))", null).ToString();
              
            }
            else if (TaskType == "纯盘")
            {
                Cur_TaskCode = dbHelper.ExecuteScalar(@"select 编号 from XW_FileOrderinfo a where 提交否='否' and 文件名!='' and not exists(select * from db_File where 编号=a.[编号] and (可读='否' or 提取='否') )", null).ToString();
            }
            string SQL = "select 文件名 from XW_FileOrderinfo where 编号='" + Cur_TaskCode + "'";
            string PDFFileName = dbHelper.ExecuteScalar(SQL, null).ToString();

            //caj控件打开文件时会出发pagechange事件,修改Cur_TaskCode为第一个任务编号,所以处理一下
            string tempTaskCode = Cur_TaskCode;
            TextLog.WritwLog("要打开的文件编号:" + Cur_TaskCode);
            axCAJAX1.Open(PDFFileName);

            //axCAJAX1.RefreshDisplay();
            axCAJAX1.Zoom(3, 1.5F);
            axCAJAX1.SetPageBrowseType(0);
            Cur_TaskCode = tempTaskCode;

            tb_Code.Text = Cur_TaskCode;
            listView2.Focus();//焦点定位到listview上   
            listView2.Items[Cur_TaskCode].Selected = true;//选中该行   
            listView2.Items[Cur_TaskCode].EnsureVisible();
            listView2.Items[Cur_TaskCode].BackColor = Color.GreenYellow;
            totalPageCount = axCAJAX1.GetPageCount();
            label_1.Text = "1/" + totalPageCount;
            label_2.Text = listView2.Items[Cur_TaskCode].Index + 1 + "/" + totalTaskCount;
            DataTable dt = dbHelper.ExecuteDataTable("select 文件名,路径,顺序,整理路径 from db_File where 编号='" + Cur_TaskCode + "'", null);
            dataGridView1.DataSource = dt;

        }

        public void Copy(FileInfo file, string desPath)
        {
            if (!desPath.EndsWith("\\"))
                desPath += "\\";
            try
            {
                file.CopyTo(desPath  + file.Name, false);
            }
            catch (IOException)
            {
                string dirname = file.DirectoryName;
                dirname=dirname.Substring(dirname.LastIndexOf('\\'));
                string newName =dirname+"--"+ file.Name;
                file.CopyTo(desPath + newName);
            }
        }

        /// <summary>
        /// 加载任务编号列表
        /// </summary>
        /// <param name="allFlag">true为全部重新加载,false为增项加载</param>
        public void loadListView()
        {
            string strSql = @"select 编号 from XW_FileOrderinfo where 文件名!='' order by 编号";
            listView2.Items.Clear();
            System.Data.DataTable dt = dbHelper.ExecuteDataTable(strSql, null);
           
            
            foreach (DataRow row in dt.Rows)
            {
                if (listView2.Items.Find(row[0].ToString(), false).Length == 1)
                    continue;
                ListViewItem lv = new ListViewItem(row[0].ToString()) { Name = row[0].ToString() };
                lv.SubItems.Add(new ListViewItem.ListViewSubItem());
                string sql = "";
                switch (TaskType)
                { 
                    case "纯盘":
                        sql = "select 提交否 from XW_FileOrderinfo where 编号='" + row[0].ToString() + "'";
                        break;
                        
                    case "刊盘":
                        sql = "select 保存否 from db_State where 编号='" + row[0].ToString() + "'";
                        break;
                }
                string result = dbHelper.ExecuteScalar(sql, null).ToString();
                if (result == "是"||result=="不可做")
                    lv.ImageIndex = 0;
                else
                    lv.ImageIndex = 1;
                listView2.Items.Add(lv);
            }
            totalTaskCount = listView2.Items.Count;
            //listView2.Items[Cur_TaskCode].BackColor = Color.GreenYellow;
        }

        /// <summary>
        /// 保存 刊盘任务信息
        /// </summary>
        private void SaveInfo(object code)
        {
            if (dbHelper.ExecuteScalar("select 保存否 from db_State where 编号='" + code + "'", null).ToString() == "不可做")
            {
                return;
            }
            listView2.Items[code.ToString()].SubItems[1].Text = "提交中";
            sHelper.Kanpan_Submit(code, GetBiaoZhu("刊盘"),TempDic,Work_Path);
            //listView2.Items[code.ToString()].ImageIndex = 0;
        }

   
        /// <summary>
        /// 复制文件夹（及文件夹下所有子文件夹和文件）
        /// </summary>
        /// <param name="sourcePath">待复制的文件夹路径</param>
        /// <param name="destinationPath">目标路径</param>
        public static void CopyDirectory(String sourcePath, String destinationPath)
        {
            DirectoryInfo info = new DirectoryInfo(sourcePath);
            Directory.CreateDirectory(destinationPath);
            foreach (FileSystemInfo fsi in info.GetFileSystemInfos())
            {
                String destName = Path.Combine(destinationPath, fsi.Name);

                if (fsi is System.IO.FileInfo)          //如果是文件，复制文件
                    File.Copy(fsi.FullName, destName);
                else                                    //如果是文件夹，新建文件夹，递归
                {
                    Directory.CreateDirectory(destName);
                    CopyDirectory(fsi.FullName, destName);
                }
            }
        }

        /// <summary>
        /// 双击任务编号
        /// </summary>
        private void listView2_DoubleClick(object sender, EventArgs e)
        {
            int selectCount = listView2.SelectedItems.Count;
            if (selectCount > 0)
            {
                string code = listView2.SelectedItems[0].SubItems[0].Text;
                //tb_Code.Text = code;

                //int page = Convert.ToInt32(dbHelper.ExecuteScalar("select min(起始页) from db_File where  编号='" + code + "' and 起始页!=0", null));
                TextLog.WritwLog("listView2_DoubleClick:"+code);
                PdfFileChangeEvent(code);
                //if (PDFFileName != pdfFileName)
                //{
                //    Handle_Flag = true;
                //    PdfFileChangeEvent(pdfFileName, page,true);
                //}
                //else
                //{
                //    Handle_Flag = false;
                //}
                //axCAJAX1.TurnToPage(page,0);
            }
        }

        /// <summary>
        /// 纯盘篇发布
        /// </summary>
        /// <param name="code">任务编号</param>
        private void Submit_Chunpan(object code)
        {
            if (dbHelper.ExecuteScalar("select 提交否 from XW_FileOrderinfo where 编号='" + code + "'", null).ToString() != "否")
            {
                return;
            }
            listView2.Items[code.ToString()].SubItems[1].Text = "提交中";
            sHelper.ChunPan_Submit(code, GetBiaoZhu("纯盘"), TempDic, Work_Path);
           
        }


        /// <summary>
        /// 纯盘设置完成标记,份提交
        /// </summary>
        private void Total_Submit_ChunPan(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定提交该任务吗?", "提交任务", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            TextLog.WritwLog("个数:"+ Traverse(Root_Path));
            if (Convert.ToInt32(dbHelper.ExecuteScalar("select count(1) from db_File", null)) < Traverse(Root_Path))
            {
                MessageBox.Show("设置完成标记失败,还有文件未整理!");
                return;
            }
            if (dbHelper.ExecuteDataTable("select 编号 from XW_FileOrderinfo where 提交否='否'", null).Rows.Count != 0)
                //if (dbHelper.ExecuteDataTable("select 编号 from XW_FileOrderinfo a where (select max(结束页) from db_File where 编号=a.编号)!=0 and 提交否='否'", null).Rows.Count != 0)
            {
                MessageBox.Show("设置完成标记失败:还有篇处理任务未提交");
                return;
            }
            Register reg = new Register();
            string Ini_Path = reg.ReadRegeditKey("QueuePath", "software\\CNKI\\Assistant\\Work\\", Register.RegDomain.CurrentUser).ToString();
            INIManage ini = new INIManage(Ini_Path);
            if (ini.SectionValues("Task") == null)
            {
                MessageBox.Show("没有任务");
                return;
            }

            MatchTask task = null;
            string strTask = "", key = "";
            foreach (string strValue in ini.SectionValues("Task"))
            {
                strTask = strValue.Substring(strValue.IndexOf('=') + 1);
                key = strValue.Substring(0, strValue.IndexOf('='));
                MatchTask mt = strTask.FromJson<MatchTask>();
                if (mt.TaskStatus == "0")
                {
                    task = mt;
                    break;
                }
            }
            task.TaskStatus = "2";
            task.FinishTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string json = task.ToJson();
            ini.setKeyValue("Task", key, json);

            MessageBox.Show("提交完成");
            if (Application.OpenForms["frmTidy"] != null)
                Application.OpenForms["frmTidy"].Close();
            this.Close();
            ReloadTask();
        }
        private void Total_Submit_KanPan(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定提交该任务吗?", "提交任务", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            if (Convert.ToInt32(dbHelper.ExecuteScalar("select count(1) from db_File", null)) < Traverse(Root_Path))
            {
                MessageBox.Show("设置完成标记失败,还有文件未整理!");
                return;
            }
            if (dbHelper.ExecuteReader("select 编号 from db_State where 保存否='否'", null).HasRows)
            {
                MessageBox.Show("还有篇任务未保存,请全部保存后在提交!");
                return;
            }

            Register reg = new Register();
            string Ini_Path = reg.ReadRegeditKey("QueuePath", "software\\CNKI\\Assistant\\Work\\", Register.RegDomain.CurrentUser).ToString();
            INIManage ini = new INIManage(Ini_Path);
            if (ini.SectionValues("Task") == null)
            {
                MessageBox.Show("没有任务");
                return;
            }

            MatchTask task = null;
            string strTask = "", key = "";
            foreach (string strValue in ini.SectionValues("Task"))
            {
                strTask = strValue.Substring(strValue.IndexOf('=') + 1);
                key = strValue.Substring(0, strValue.IndexOf('='));
                MatchTask mt = strTask.FromJson<MatchTask>();
                if (mt.TaskStatus == "0")
                {
                    task = mt;
                    break;
                }
            }
            task.TaskStatus = "2";
            task.FinishTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string json = task.ToJson();
            ini.setKeyValue("Task", key, json);

            MessageBox.Show("提交完成");
            if (Application.OpenForms["frmTidy"] != null)
                Application.OpenForms["frmTidy"].Close();
            this.Close();
            ConfigHelper.SetValue("BEGIN_CODE_Kanpan_" + GongHao, null);
            ReloadTask();
        }
       

        /// <summary>
        /// 获取用户标注信息
        /// </summary>
        /// <param name="tasktype">纯盘or刊盘</param>
        /// <returns></returns>
        private Dictionary<string, object> GetBiaoZhu(string tasktype)
        {
            Dictionary<string, object> dic = new Dictionary<string, object>();
            dic.Add("保密", cb_secret.Checked ? "是" : "否");
            dic.Add("删除字样", cb_delete.Checked ? "是" : "否");
            dic.Add("滞后上网", cb_DelayDate.Checked ? "是" : "否");
            if (tasktype == "刊盘")
            {
                DataTable dt_text = dbHelper.ExecuteDataTable("select 摘要 from db_File where 编号='" + tb_Code.Text + "'", null);
                string text = "";
                foreach (DataRow row in dt_text.Rows)
                {
                    text += row[0].ToString() + " ";
                }
                text = text.Replace("\"", "");
                dic.Add("摘要", text);
                //string sql = "update XW_FileOrderinfo set 保密否='" + dic["保密"].ToString() + "',删除字样='" + dic["删除字样"] + "',论文摘要=\"" + text + "\" where 编号='" + code + "'";
            }
            else if (tasktype == "纯盘")
            {
                string ShouQuanFanKui = string.Empty;
                if (radioButton5.Checked)
                {
                    dic.Add("备注", null);
                    ShouQuanFanKui = "是";
                }
                else if (radioButton6.Checked)
                {
                    dic.Add("备注", cb_explain.Text);
                    ShouQuanFanKui = "否";
                }

                else if (radioButton7.Checked)
                {
                    dic.Add("备注", null);
                    ShouQuanFanKui = "不合格";
                }
                dic.Add("版权反馈", ShouQuanFanKui);
                dic.Add("授权", cb_authoration.Checked ? "是" : "否");//反义一下, 由于后续岗位的列名是'无授权','无作者签名'
                dic.Add("签名", cb_signature.Checked ? "是" : "否");  //反义一下, 由于后续岗位的列名是'无授权','无作者签名'
                if (cb_year.Checked)
                    dic.Add("学位年度", null);
                else
                    dic.Add("学位年度", num_year.Value);

                if (radioButton1.Checked)
                    dic.Add("级别", "硕士");
                else if (radioButton2.Checked)
                    dic.Add("级别", "博士");
                else if (radioButton3.Checked)
                    dic.Add("级别", "博士后");
                else if (radioButton4.Checked)
                    dic.Add("级别", null);
            }


            return dic;
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked)
            {
                cb_explain.Enabled = true;
                cb_explain.SelectedIndex = 0;
                radioButton6.ForeColor = Color.Red;
            }
            else
            {
                cb_explain.Enabled = false;
                radioButton6.ForeColor = Color.Black;
            }
        }

        private void cb_year_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_year.Checked)
                num_year.Enabled = false;
            else
                num_year.Enabled = true;
        }


        /// <summary>
        /// pdf文件翻页事件
        /// </summary>
        private void axCAJAX1_PageChanged(object sender, AxCAJAXLib._DCAJAXEvents_PageChangedEvent e)
        {
            label_1.Text = e.index + "/" + totalPageCount;
        }

        /// <summary>
        /// 刊盘 修改小样编号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellEndEdit_Kanpan(object sender, DataGridViewCellEventArgs e)
        {
            string newName = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            string path = dataGridView1.Rows[e.RowIndex].Cells["路径"].Value.ToString();
            string Cur_TaskCode = tb_Code.Text;
            try
            {
                int num = Convert.ToInt32(newName);
                if (num < 1 || num > 99)
                {
                    MessageBox.Show("小样编号必须是小于100的正数", "Error");
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dbHelper.ExecuteScalar("select 顺序 from  db_File  where 路径='" + path + "'", null);
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("小样编号必须是数字", "Error");
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dbHelper.ExecuteScalar("select 顺序 from  db_File  where 路径='" + path + "'", null);
                return;
            }
            string afterPath = dataGridView1.Rows[e.RowIndex].Cells["整理路径"].Value.ToString();
            TextLog.WritwLog("整理路径:" + afterPath);
            FileInfo oldFile = new FileInfo(afterPath);
            string newPath = Work_Path + "\\整理后\\" + Cur_TaskCode + "\\" + newName + "_" + oldFile.Name;
            TextLog.WritwLog("整理后路径:" + newPath);
            oldFile.MoveTo(newPath);
            TextLog.WritwLog("移动成功");
            //string path = dataGridView1.Rows[e.RowIndex].Cells["路径"].Value.ToString();
            dbHelper.ExecuteNonQuery("update db_File set 整理路径='" + newPath + "', 顺序='" + newName + "' where 路径='" + path + "'", null);
            dataGridView1.Rows[e.RowIndex].Cells["整理路径"].Value = newPath;

        }

        private void dataGridView1_CellEndEdit_Chunpan(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                
                string path = dataGridView1.Rows[e.RowIndex].Cells["路径"].Value.ToString();
                TextLog.WritwLog("path="+path);
                string Cur_TaskCode = tb_Code.Text;
                string sql = TaskType == "纯盘" ? "select 提交否 from XW_FileOrderinfo where 编号=(select 编号 from  db_File  where 路径='" + path + "')" :
                    "select 保存否 from db_State where 编号=(select 编号 from  db_File  where 路径='" + path + "')";
                if (dbHelper.ExecuteScalar(sql, null).ToString() == "是")
                {
                    MessageBox.Show("文件已上传提交,无法标注顺序");
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dbHelper.ExecuteScalar("select 顺序 from  db_File  where 路径='" + path + "'", null);
                    return;
                }
                string newName = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                TextLog.WritwLog("newName=" + newName);
                try
                {
                    int num = Convert.ToInt32(newName);
                    if (num < 1 || num > 99)
                    {
                        MessageBox.Show("小样编号必须是小于100的正数", "Error");
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dbHelper.ExecuteScalar("select 顺序 from  db_File  where 路径='" + path + "'", null);
                        return;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("小样编号必须是数字", "Error");
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dbHelper.ExecuteScalar("select 顺序 from  db_File  where 路径='" + path + "'", null);
                    return;
                }
                TextLog.WritwLog("f3");
                string afterPath = dataGridView1.Rows[e.RowIndex].Cells["整理路径"].Value.ToString();
                string newPath;
                if (File.Exists(afterPath))
                {
                    FileInfo oldFile = new FileInfo(afterPath);
                    newPath = oldFile.Directory.FullName+"\\"+ newName + "_" + oldFile.Name;
                    oldFile.MoveTo(newPath);
                }
                else
                { 
                    string Cur_filename=afterPath.Substring(afterPath.LastIndexOf('\\')+1);
                    FileInfo file=new FileInfo(Work_Path+"\\ArticleUpload\\"+Cur_TaskCode+"\\"+Cur_filename);
                    newPath =file.Directory.FullName+ "\\" +newName+"_"+ file.Name;
                    file.MoveTo(newPath);
                }
                TextLog.WritwLog("f4");
                dbHelper.ExecuteNonQuery("update db_File set 整理路径='" + newPath + "',顺序='" + newName + "' where 路径='" + path + "'", null);
                dataGridView1.Rows[e.RowIndex].Cells["整理路径"].Value = newPath;
            }
            catch (InvalidExpressionException sq)
            {
                MessageBox.Show(sq.Message);
            }
        }

        /// <summary>
        /// 异步接收udp消息
        /// </summary>
        public void UdpReceive()
        {
                PublicTool.localUdp.BeginReceive(new AsyncCallback(FinalReceiveCallback), null);
        }
        static object obj=new object();
        public void FinalReceiveCallback(IAsyncResult iar)
        {
            lock (obj)
            {
            TextLog.WritwLog("callback");
            try
            {
                if (iar.IsCompleted)
                {
                    Byte[] receiveBytes = PublicTool.localUdp.EndReceive(iar, ref remoteEP);
                    string FinalBackString = Encoding.GetEncoding("GB2312").GetString(receiveBytes);
                    TextLog.WritwLog("callback==" + FinalBackString);
                    #region 处理过程
                    try
                    {
                        if (FinalBackString != "Y")
                        {
                            SubmitResult sr = FinalBackString.FromJson<SubmitResult>();
                            string code = sr.ArticleCode;
                            TextLog.WritwLog(code + ":" + FinalBackString);
                            TextLog.WritwLog("返回:" + sr.ArticleCode + "," + sr.Statu + "," + sr.ErrInfo);
                            if (sr.Statu.ToUpper() == "TRUE")
                            {
                                TextLog.WritwLog("返回编码:" + code);
                                string dics="";
                                foreach (string key in TempDic.Keys)
                                {
                                    dics += key + "-";
                                }
                                TextLog.WritwLog("删除前字典里:"+dics);
                                try{
                                    TextLog.WritwLog("isRead:"+TempDic[code]);
                                }
                                catch(Exception e)
                                {
                                    TextLog.WritwLog("取字典报错:"+e.Message);
                                }
                                string isRead = TempDic[code];
                                this.BeginInvoke(new Action(() =>
                                {
                                    TextLog.WritwLog("TaskType:" + TaskType);
                                    //string sql = string.Format("update XW_FileOrderinfo set 年度='{0}',级别='{1}',保密否='{2}',版权反馈否='{3}',是否签名='{4}',是否授权='{5}',备注='{6}',删除字样='{7}',提交否='是' where 编号='{8}'",
                                    //   sTask.Year, sTask.Level, sTask.IsSecret, sTask.Iscopyright, sTask.IsQM, sTask.IsSQ, sTask.Explain, sTask.DeleteWords, code);
                                    if(TaskType=="纯盘")
                                    {
                                        TextLog.WritwLog("callback纯盘.isRead="+isRead);
                                        switch (isRead)
                                        { 
                                            case "是":
                                                try
                                                {
                                                    TextLog.WritwLog("update XW_FileOrderinfo set 提交否='是' where 编号='" + code + "'");
                                                    dbHelper.ExecuteNonQuery("update XW_FileOrderinfo set 提交否='是' where 编号='" + code + "'", null);
                                                }
                                                catch (Exception sfs)
                                                {
                                                    TextLog.WritwLog(sfs.Message);
                                                }
                                                break;
                                            case "否":
                                                TextLog.WritwLog("update XW_FileOrderinfo set 提交否='不可做' where 编号='" + code + "'");
                                                dbHelper.ExecuteNonQuery("update XW_FileOrderinfo set 提交否='不可做' where 编号='" + code + "'", null);
                                                break;
                                        }
                                        
                                    }
                                    else if (TaskType == "刊盘")
                                    {
                                        TextLog.WritwLog("callback刊盘");
                                        switch (isRead)
                                        {
                                            case "是":
                                                TextLog.WritwLog("update db_State set 保存否='是' where 编号='" + code + "'");
                                                dbHelper.ExecuteNonQuery("update db_State set 保存否='是' where 编号='" + code + "'", null);
                                                break;
                                            case "否":
                                                dbHelper.ExecuteNonQuery("update db_State set 保存否='不可做' where 编号='" + code + "'", null);
                                                dbHelper.ExecuteNonQuery("update XW_FileOrderinfo set 可读否='否' where 编号='" + code + "'", null);
                                                break;
                                        }
                                    }
                                    TextLog.WritwLog("更新完db");
                                    listView2.Items[code.ToString()].ImageIndex = 0;
                                    listView2.Items[code.ToString()].SubItems[1].Text = "";
                                    TextLog.WritwLog("删除编号:" + code);
                                    TempDic.Remove(code);
                                    string ds = "";
                                    foreach (string key in TempDic.Keys)
                                    {
                                        ds += key + "-";
                                    }
                                    TextLog.WritwLog("删除后字典里:" + ds);

                                }));
                            }
                            else
                            {
                                this.BeginInvoke(new Action(() =>
                                {
                                    listView2.Items[code].SubItems[1].Text = "提交失败";
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
                        this.BeginInvoke(new Action(() =>
                                {
                                    listView2.Items[code].SubItems[1].Text = "提交失败";
                                }));
                    }
                    #endregion
                    PublicTool.localUdp.BeginReceive(new AsyncCallback(FinalReceiveCallback), null);
                }
            }
            catch (Exception e)
            {
                TextLog.WritwLog(e.Message);
            }
            }
        }

        /// <summary>
        /// 切换任务后更新界面
        /// </summary>
        /// <param name="code">任务编号</param>
        public void UpdateUi(string code)
        {
            try
            {
                string pdfFileName = dbHelper.ExecuteScalar("select 文件名 from XW_FileOrderinfo where  编号='" + code + "'", null).ToString();
                
                axCAJAX1.Open(pdfFileName);
                axCAJAX1.Zoom(3, 1.5F);
                //axCAJAX1.SetPageBrowseType(0);
                //axCAJAX1.RefreshDisplay();
                totalPageCount = axCAJAX1.GetPageCount();
                label_2.Text = listView2.Items[code].Index + 1 + "/" + listView2.Items.Count;
                //更新任务列表任务项北京
                try
                {
                    listView2.Items[tb_Code.Text].BackColor = Color.White;
                }
                catch (Exception e) {
                    TextLog.WritwLog("UpdateUi(更新backcolor):" + e.Message);
                }

                tb_Code.Text = code;

                listView2.Items[code].BackColor = Color.GreenYellow;
                listView2.Items[code].EnsureVisible();
                //更新任务包含的文件列表
                DataTable dt = dbHelper.ExecuteDataTable("select 文件名,路径,顺序,整理路径 from db_File where 编号='" + code + "'", null);
                dataGridView1.DataSource = dt;

                //更新右侧标注结果
                DataRow row = dbHelper.ExecuteDataTable("select * from XW_FileOrderinfo where 编号='" + code + "'", null).Rows[0];
                //备注是null则代表该任务还没有提交保存过,所以默认界面为上一任务的界面,
                //如果备注不是null,则该任务提交成功或失败,界面应显示用户保存了的参数
                if (row["备注"] == DBNull.Value)
                    return;
                if (TaskType == "纯盘")
                {
                    if (row["是否授权"] == DBNull.Value || row["是否授权"].ToString() == "否")
                        cb_authoration.Checked = false;
                    else if (row["是否授权"].ToString() == "是")
                        cb_authoration.Checked = true;
                    if (row["是否签名"] == DBNull.Value || row["是否签名"].ToString() == "否")
                        cb_signature.Checked = false;
                    else if (row["是否签名"].ToString() == "是")
                        cb_signature.Checked = true;

                    if (row["年度"] == DBNull.Value || row["年度"].ToString() == "")
                        cb_year.Checked = true;
                    else
                    {
                        cb_year.Checked = false;
                        num_year.Value = Convert.ToDecimal(row["年度"]);
                    }
                    if (row["级别"] == DBNull.Value || row["级别"].ToString() == "")
                        radioButton4.Checked = true;
                    else if (row["级别"].ToString() == "硕士")
                        radioButton1.Checked = true;
                    else if (row["级别"].ToString() == "博士")
                        radioButton2.Checked = true;
                    else if (row["级别"].ToString() == "博士后")
                        radioButton3.Checked = true;
                }
                if (row["滞后上网"] == DBNull.Value || row["滞后上网"].ToString() == "否")
                    cb_DelayDate.Checked = false;
                else
                    cb_DelayDate.Checked = true;
                if (row["保密否"] == DBNull.Value || row["保密否"].ToString() == "否")
                    cb_secret.Checked = false;
                else
                    cb_secret.Checked = true;
                if (row["删除字样"] == DBNull.Value || row["删除字样"].ToString() == "否")
                    cb_delete.Checked = false;
                else
                    cb_delete.Checked = true;
                if (row["版权反馈否"] == DBNull.Value || row["版权反馈否"].ToString() == "否")
                    radioButton6.Checked = true;
                else if (row["版权反馈否"].ToString() == "是")
                    radioButton5.Checked = true;
                else if (row["版权反馈否"].ToString() == "不合格")
                    radioButton7.Checked = true;
                if (row["备注"] == DBNull.Value)
                    cb_explain.SelectedText = null;
                else
                {
                    cb_explain.Text = row["备注"].ToString();
                }
            }
            catch (Exception es)
            {
                TextLog.WritwLog("UpdateUi:"+es.Message);
            }
        }

        /// <summary>
        /// 篇提交
        /// </summary>
        /// <param name="code">要提交的任务编号</param>
        public void CodeChangeSubmit(string code)
        {
            //string code = label_2.Text;
            //这里为防止正在提叫但还未返回时任务重复提交
            string state = listView2.Items[code].SubItems[1].Text;
            if (state == "提交中" || state == "置不可做中")
                return;
            PianSubmit(code);
        }

        /// <summary>
        /// 构造队列存储页码, 用于检测何时切换pdf
        /// </summary>
        private System.Collections.Queue ScrollPosQueue = new System.Collections.Queue();
        private void axCAJAX1_MouseWheelEvent(object sender, AxCAJAXLib._DCAJAXEvents_MouseWheelEvent e)
        {
            ScrollPosQueue.Enqueue(axCAJAX1.GetScrollPos(1));
            if (ScrollPosQueue.Count > 10)
            {
                ScrollPosQueue.Dequeue();
                int i = Convert.ToInt32(ScrollPosQueue.Peek());
                if (i == 0)
                    return;
                foreach (var s in ScrollPosQueue)
                {
                    if (Convert.ToInt32(s) != i)
                        return;

                }
                
                ChangeNextPDF();
            }
        }

        public void ChangeNextPDF()
        {
            string code = tb_Code.Text;
            if (!SubmitCheckEvent(code))
                return;
            CodeChangeSubmit(code);
            int index=listView2.Items[code].Index;
            if (index + 1 < listView2.Items.Count)
            {
                index++;
                code = listView2.Items[index].Text;
                PdfFileChangeEvent(code);
            }
        }

        private void frmTidy_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (ListViewItem item in listView2.Items)
            {
                if (item.SubItems.Count > 1)
                {
                    if (item.SubItems[1].Text == "提交中")
                    {
                        e.Cancel = true;
                        MessageBox.Show("有任务正在提交,无法关闭");
                        return;
                    }
                }
            }
        }

        private void frmTidy_FormClosed(object sender, FormClosedEventArgs e)
        {
           // UpdateMainWindow();
            UnregisterHotKey(this.Handle, 888);
            UnregisterHotKey(this.Handle, 777);
            backgroudThread.Abort();
        }

        private void cb_authoration_CheckStateChanged(object sender, EventArgs e)
        {
            if (cb_authoration.Checked)
            {
                //groupBox5.Enabled = false;
                radioButton5.Checked = true;
                cb_signature.Enabled = false;
                cb_signature.Checked = false;
                radioButton5.Enabled = radioButton6.Enabled = radioButton7.Enabled = cb_explain.Enabled = false;
                cb_authoration.ForeColor = Color.Red;
            }
            else
            {
                cb_signature.Enabled = radioButton5.Enabled = radioButton6.Enabled = radioButton7.Enabled = cb_explain.Enabled = true;
                cb_authoration.ForeColor = Color.Black;
            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked)
            {
                cb_signature.Enabled = false;
                cb_signature.Checked = false;
                radioButton7.ForeColor = Color.Red;
            }
            else
            {
                cb_signature.Enabled = true;
                radioButton7.ForeColor = Color.Black;
            }
        }

        /// <summary>
        /// 置不可做
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            FrmNotDo fn = new FrmNotDo();
            fn.StartPosition = FormStartPosition.CenterParent;
            if (fn.ShowDialog() == DialogResult.OK)
            {
               
                switch (TaskType)
                {
                    case "纯盘":
                        dbHelper.ExecuteNonQuery("update XW_FileOrderinfo set 提交否='不可做' where 编号='" + tb_Code.Text + "'", null);
                        listView2.Items[tb_Code.Text].SubItems[1].Text = "置不可做中";
                        sHelper.ChunPan_Submit_UnRead(tb_Code.Text, fn.info, Work_Path, TempDic);
                        break;
                    case "刊盘":
                        string sql = "update XW_FileOrderinfo set 可读否='否',制作说明=\"" + fn.info + "\" where 编号='" + tb_Code.Text + "'";
                        dbHelper.ExecuteNonQuery(sql, null);
                        sql = "update db_State set 保存否='不可做' where 编号='" + tb_Code.Text + "'";
                        dbHelper.ExecuteNonQuery(sql, null);
                        listView2.Items[tb_Code.Text].SubItems[1].Text = "置不可做中";
                        sHelper.KanPan_Submit_UnRead(tb_Code.Text, fn.info,Work_Path,TempDic);
                        //listView2.Items[tb_Code.Text].ImageIndex = 0;
                        break;
                }
            }
        }

        private void cb_year_CheckedChanged_1(object sender, EventArgs e)
        {
            num_year.Enabled = !cb_year.Checked;
            if (cb_year.Checked)
                cb_year.ForeColor = Color.Red;
            else
                cb_year.ForeColor = Color.Black;
        }

        /// <summary>
        /// 计算目录中所有嵌套的文件的个数
        /// </summary>
        /// <param name="sPathName">文件夹路径</param>
        /// <returns>文件个数</returns>
        public int Traverse(string sPathName)
        {
            //创建一个队列用于保存子目录
            int i = 0;
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
                    if (fi.Name.StartsWith("~$")||fi.Extension==".ini")
                        continue;
                    i++;
                }
            }
            return i;
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            System.Diagnostics.Process.Start(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
        }

        /// <summary>
        /// 检查小样是否全部都编号,检查年份是否符合规则
        /// </summary>
        /// <param name="code">任务编号</param>
        /// <param name="page">如果返回false,则返回原来的页数</param>
        /// <returns>全部编号返回true,否则返回false</returns>
        private bool ChunPanCheck(string code)
        {
            if (num_year.Value > 2080 || num_year.Value < 1900)
            {
                UnregisterHotKey(this.Handle, 888);
                UnregisterHotKey(this.Handle, 777);
                MessageBox.Show("学位年度输入超范围", "Error");
                RegisterHotKey(this.Handle, 888, 0, Keys.Down);
                RegisterHotKey(this.Handle, 777, 0, Keys.Up);
                return false;
            }
            if (dataGridView1.Rows.Count > 1)
            {
                string sql = "select 顺序 from db_File where 编号=" + code;
                DataTable dt = dbHelper.ExecuteDataTable(sql, null);
                foreach (DataRow row in dt.Rows)
                {
                    if (row[0].ToString() == "")
                    {
                        UnregisterHotKey(this.Handle, 888);
                        UnregisterHotKey(this.Handle, 777);
                        MessageBox.Show("存在未编号小样", "Error");
                        RegisterHotKey(this.Handle, 888, 0, Keys.Down);
                        RegisterHotKey(this.Handle, 777, 0, Keys.Up);
                        return false;
                    }
                }
                try
                {
                    dt.PrimaryKey = new DataColumn[] { dt.Columns[0] };
                }
                catch (Exception e)
                {
                    UnregisterHotKey(this.Handle, 888);
                    UnregisterHotKey(this.Handle, 777);
                    MessageBox.Show("小样编号不能重复","Error");
                    RegisterHotKey(this.Handle, 888, 0, Keys.Down);
                    RegisterHotKey(this.Handle, 777, 0, Keys.Up);
                    return false;
                }
                return true;
            }
            return true;
        }

        private bool KanPanCheck(string code)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                string sql = "select 顺序 from db_File where 编号='" + code+"'";
                DataTable dt = dbHelper.ExecuteDataTable(sql, null);
                foreach (DataRow row in dt.Rows)
                {
                    if (row[0].ToString() == "")
                    {
                        UnregisterHotKey(this.Handle, 888);
                        UnregisterHotKey(this.Handle, 777);
                        MessageBox.Show("存在未编号小样", "Error");
                        //axCAJAX1.TurnToPage(page, 0);
                        RegisterHotKey(this.Handle, 888, 0, Keys.Down);
                        RegisterHotKey(this.Handle, 777, 0, Keys.Up);
                        return false;
                    }
                }
                try
                {
                    dt.PrimaryKey = new DataColumn[] { dt.Columns[0] };
                }
                catch (Exception e)
                {
                    UnregisterHotKey(this.Handle, 888);
                    UnregisterHotKey(this.Handle, 777);
                    MessageBox.Show("小样编号不能重复", "Error");
                    //axCAJAX1.TurnToPage(page, 0);
                    RegisterHotKey(this.Handle, 888, 0, Keys.Down);
                    RegisterHotKey(this.Handle, 777, 0, Keys.Up);
                    return false;
                }
                return true;
            }
            return true;
        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 0 || e.RowIndex < 0 || dataGridView1.Rows.Count <= 0) return;
            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value ?? string.Empty).ToString();
        }

        public void DeleteXiaoYang()
        {
            if (dataGridView1.Focused)
            {
                if (dataGridView1.SelectedRows != null)
                {
                    if (MessageBox.Show("确定删除文件:" + dataGridView1.CurrentRow.Cells[0].Value + "?", "警告", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        string path = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                        File.Delete(path);
                        dbHelper.ExecuteNonQuery("delete from db_File where 路径 = '" + path + "'", null);
                    }
                }
            }
        }
        /// <summary>
        /// 任务存疑
        /// </summary>
        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定把该任务置疑吗?", "任务存疑", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            Register reg = new Register();
            string Ini_Path = reg.ReadRegeditKey("QueuePath", "software\\CNKI\\Assistant\\Work\\", Register.RegDomain.CurrentUser).ToString();
            INIManage ini = new INIManage(Ini_Path);
            if (ini.SectionValues("Task") == null)
            {
                MessageBox.Show("没有任务");
                return;
            }

            MatchTask task = null;
            string strTask = "", key = "";
            foreach (string strValue in ini.SectionValues("Task"))
            {
                strTask = strValue.Substring(strValue.IndexOf('=') + 1);
                key = strValue.Substring(0, strValue.IndexOf('='));
                MatchTask mt = strTask.FromJson<MatchTask>();
                if (mt.TaskStatus == "0")
                {
                    task = mt;
                    break;
                }
            }
            task.TaskStatus = "8";
            task.FinishTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string json = task.ToJson();
            ini.setKeyValue("Task", key, json);

            MessageBox.Show("任务已置疑");
            if (Application.OpenForms["frmTidy"] != null)
                Application.OpenForms["frmTidy"].Close();
            this.Close();
            ReloadTask();
        }

        private void checkOrRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            
            try
            {
                CheckBox button = sender as CheckBox;
                if (button.Checked)
                    button.ForeColor = Color.Red;
                else
                    button.ForeColor = Color.Black;
            }
            catch
            {
                RadioButton rb = sender as RadioButton;
                if (rb.Checked)
                    rb.ForeColor = Color.Red;
                else
                    rb.ForeColor = Color.Black;

            }
        }
       
       

       
    }

}
