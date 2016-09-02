using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Manual_Import.Model;
using Manual_Import;
using Manual_Import.Helper;
using System.Xml.Linq;
using PDF = iTextSharp.text;
using iTextSharp.text.pdf;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Utility.Common;
using System.Runtime.InteropServices;
using SharpCompress.Reader;
using SharpCompress.Common;
using System.Threading;
using Utility.Dao;
using iTextSharp.text;
using System.Threading.Tasks;
using iTextSharp.text.pdf.parser;

namespace Manual_Import.ViewModel
{
    public class TidyHelper
    {
        object G_missing = System.Reflection.Missing.Value;
        object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
        object What = Word.WdGoToItem.wdGoToSection;
        object Which = Word.WdGoToDirection.wdGoToFirst;

        ViewModel_Main ViewModel;
        Action<string, int> ShowLog;
        Action<Model_FileSystem> RemoveItem;
        public Action DoMany;//批量生成
        public Action DoAlone;//单独生成
        public Action DoError;//错误文件重新整理
        SQLiteDBHelper dbHelper;
        public CancellationTokenSource tokenSource;

        string[] Front5Array = { };
        string[] Back5Array = { };
        const string TEMP_PDF = "nothing.pdf";
        public int PDF_FRONT_NUM = Convert.ToInt32(ConfigHelper.GetValue("PDF_FrontNum"));
        public int PDF_BACK_NUM = Convert.ToInt32(ConfigHelper.GetValue("PDF_BackNum"));
        public int WORD_FRONT_NUM = Convert.ToInt32(ConfigHelper.GetValue("Word_FrontNum"));
        public int WORD_BACK_NUM = Convert.ToInt32(ConfigHelper.GetValue("Word_BackNum"));
        public bool CUSTOM_DEFINE = false;

        //public string CurPdfName;//当前合成pdf文件名,在外赋值

        public TidyHelper(SQLiteDBHelper sqlite, ViewModel_Main Vmodel, Action<string, int> Show, Action<Model_FileSystem> removeItem)
        {
            dbHelper = sqlite;
            ViewModel = Vmodel;
            ShowLog = Show;
            RemoveItem = removeItem;

            Front5Array = ConfigHelper.GetValue("Front5RemoveWord").Split(new char[] { ',' });
            Back5Array = ConfigHelper.GetValue("Back5FetchWord").Split(new char[] { ',' });

            if (Vmodel.TaskType == "纯盘")
            {
                DoAlone = ChunPan_Alone;
                DoMany = ChunPan_Many;
                DoError = ChunPan_Error;
            }
            else if (Vmodel.TaskType == "刊盘")
            {
                DoAlone = KanPan_Alone;
                DoMany = KanPan_Many;
                DoError = KanPan_Error;
            }
        }

        /// <summary>
        /// 给pdf都第一页和最后一页加水印
        /// </summary>
        /// <param name="inputfilepath"></param>
        /// <param name="pageSum"></param>
        /// <param name="code"></param>
        public void WaterMarkPdf(ref string inputfilepath, out int pageSum, string code = "")
        {
            Utility.Log.TextLog.WritwLog("开始pdf加水印");
            pageSum = 0;
            string ChangeFilePath = ViewModel.Upload_Path + "\\非正文页.pdf";
            File.Move(inputfilepath, ChangeFilePath);
            Utility.Log.TextLog.WritwLog("pdf重命名");
            if (code == "")
                inputfilepath = ViewModel.Upload_Path + "\\" + ViewModel.Begin_Code + "[整].pdf";
            else
                inputfilepath = ViewModel.Upload_Path + "\\" + code + "[整].pdf";
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(ChangeFilePath);
                Utility.Log.TextLog.WritwLog("1");
                pdfStamper = new PdfStamper(pdfReader, new FileStream(inputfilepath, FileMode.Create));
                //int total = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\SIMFANG.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState();
                Utility.Log.TextLog.WritwLog("2");

                content = pdfStamper.GetOverContent(1);//在内容上方加水印
                //透明度
                gs.FillOpacity = 0.8f;
                content.SetGState(gs);
                //content.SetGrayFill(0.3f);
                //开始写入文本
                content.BeginText();
                content.SetColorFill(BaseColor.RED);
                content.SetFontAndSize(font, 50);
                content.SetTextMatrix(0, 0);
                if (code == "")
                    content.ShowTextAligned(Element.ALIGN_CENTER, ViewModel.Begin_Code + "任务开始", width / 2, height - 50, 0);
                else
                    content.ShowTextAligned(Element.ALIGN_CENTER, code + "任务开始", width / 2, height - 50, 0);
                content.EndText();

                content = pdfStamper.GetOverContent(pdfReader.NumberOfPages);//在内容上方加水印
                //透明度
                gs.FillOpacity = 0.8f;
                content.SetGState(gs);
                //content.SetGrayFill(0.3f);
                //开始写入文本
                content.BeginText();
                content.SetColorFill(BaseColor.RED);
                content.SetFontAndSize(font, 50);
                content.SetTextMatrix(0, 0);
                if (code == "")
                    content.ShowTextAligned(Element.ALIGN_CENTER, ViewModel.Begin_Code + "任务结束", width / 2, 20, 0);
                else
                    content.ShowTextAligned(Element.ALIGN_CENTER, code + "任务结束", width / 2, 20, 0);
                content.EndText();

            }
            catch (Exception ex)
            {
                Utility.Log.TextLog.WritwLog("WaterMarkPdf:" + ex.Message);
                throw ex;
            }
            finally
            {
                pageSum = pdfReader.NumberOfPages;
                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();

                File.Delete(ChangeFilePath);
            }
        }
        /// <summary>
        /// 纯盘单独生成
        /// </summary>
        private void ChunPan_Alone()
        {
            try
            {
                ShowLog("开始生成任务", 2);
                int XiaoYangSum = 0;//小样个数
                var checkedItem = ViewModel.Models.Where(m => { return m.Checked == true && m.Name != ".."; });
                if (checkedItem.Count() == 0)
                {
                    ShowLog("没有选择任何文件", 3);
                    return;
                }
                string desDirName = ViewModel.AfterTydyPath + "\\" + ViewModel.Begin_Code;
                if (!Directory.Exists(desDirName))
                {
                    Directory.CreateDirectory(desDirName);
                }
                List<string> backPdfs = new List<string>();

                foreach (Model_FileSystem fs in checkedItem)
                {

                    #region 整理文件夹
                    if (fs.Type == SystemType.Dir)
                    {
                        //循环该文件夹中包括子文件夹中的所有文件
                        foreach (FileInfo file in Traverse(fs.FullPath))
                        {
                            string backpdf;
                            //整理纯盘文件
                            string[] arrayStr = FileOprate_Chunpan(file, out backpdf);
                            string afterCopyFilename = CopyFile(file, desDirName);
                            //每整理一个文件向表中插入一条记录
                            string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}','{4}',{5},\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, afterCopyFilename, file.Extension, arrayStr[1], arrayStr[0], 0, file.FullName, 0, "null", desDirName + "\\" + afterCopyFilename);
                            dbHelper.ExecuteNonQuery(sql, null);
                            //按整理结果更新工具界面
                            if (arrayStr[0] != "是" || arrayStr[1] != "是")
                            {
                                ViewModel.CurFail++;
                                ViewModel.TotalFail++;
                            }
                            else
                            {
                                ViewModel.CurSuc++;
                                ViewModel.TotalSuc++;
                            }
                            ViewModel.UnTidy--;

                            backPdfs.Add(backpdf);
                            XiaoYangSum++;
                        }
                    }
                    #endregion

                    #region 整理单个文件
                    else
                    {
                        FileInfo file = new FileInfo(fs.FullPath);
                        string backpdf;
                        //整理纯盘文件
                        string[] arrayStr = FileOprate_Chunpan(file, out backpdf);
                        string afterCopyFilename = CopyFile(file, desDirName);
                        //每整理一个文件向表中插入一条记录
                        string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}','{4}',{5},\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, afterCopyFilename, file.Extension, arrayStr[1], arrayStr[0], 0, file.FullName, 0, "null", desDirName + "\\" + afterCopyFilename);
                        dbHelper.ExecuteNonQuery(sql, null);
                        backPdfs.Add(backpdf);
                        //按整理结果更新工具界面
                        if (arrayStr[0] != "是" || arrayStr[1] != "是")
                        {
                            fs.HasTidy = -1;
                            ViewModel.CurFail++;
                            ViewModel.TotalFail++;
                        }
                        else
                        {
                            ViewModel.CurSuc++;
                            ViewModel.TotalSuc++;
                            fs.HasTidy = 1;
                        }
                        ViewModel.UnTidy--;
                        XiaoYangSum++;
                    }
                    #endregion
                    fs.Checked = false;
                }

                string outPdfName;
                if (backPdfs.Contains(""))
                    outPdfName = "";
                else if (backPdfs.All(a => { return a == null; }))//没有一个文件提取出有效信息,则不必合并pdf了
                {
                    if (!File.Exists(TEMP_PDF))
                        Voidpdf();
                    outPdfName = TEMP_PDF;
                }
                else
                {
                    outPdfName = ViewModel.Upload_Path + "\\" + ViewModel.Begin_Code + "[整].pdf";
                    //合并所有以上整理好的pdf
                    CombineMultiplePDFs(backPdfs, ref outPdfName);
                }
                int TotalPage = 0;
                //给pdf加水印
                if (outPdfName != "")
                    WaterMarkPdf(ref outPdfName, out TotalPage);
                string str = string.Format("insert into XW_FileOrderinfo(编号,路径,提交否,文件名,年度,级别,版权反馈否,保密否,是否签名,是否授权,删除字样,提取页数,小样数) values('{0}','{1}','否','{2}','{3}','硕士','是','否','否','否','否',{4},{5})", ViewModel.Begin_Code, desDirName, outPdfName, DateTime.Now.Year, TotalPage, XiaoYangSum);
                dbHelper.ExecuteNonQuery(str, null);

                ShowLog(ViewModel.Begin_Code + ":生成任务完毕!", 1);
                ViewModel.Begin_Code = (Convert.ToInt32(ViewModel.Begin_Code) + 1).ToString();
                //更新配置文件中的编号
                ConfigHelper.SetValue("BEGIN_CODE", ViewModel.Begin_Code);

            }
            catch (Exception ee)
            {
                ShowLog(ee.Message, 3);
                Utility.Log.TextLog.WritwLog(ee.Message);
            }
        }
        /// <summary>
        /// 纯盘批量生成
        /// </summary>
        private void ChunPan_Many()
        {
            try
            {

                ShowLog("开始生成任务", 2);
                var checkedItem = ViewModel.Models.Where(m => { return m.Checked == true && m.Name != ".."; });
                if (checkedItem.Count() == 0)
                {
                    ShowLog("没有选择任何文件", 3);
                    return;
                }

                int year = DateTime.Now.Year;
                foreach (Model_FileSystem fs in checkedItem)
                {
                    if (tokenSource.Token.IsCancellationRequested)
                    {
                        ShowLog("整理停止", 2);
                        return;
                    }
                    int XiaoYangSum = 0;


                    string desDirName = ViewModel.AfterTydyPath + "\\" + ViewModel.Begin_Code;
                    if (!Directory.Exists(desDirName))
                        Directory.CreateDirectory(desDirName);
                    string backpdf;

                    #region 勾选项是文件夹
                    if (fs.Type == SystemType.Dir)
                    {
                        var fileList = Traverse(fs.FullPath);
                        if (fileList.Where(f => { return (f.Extension.ToLower() == ".pdf" || f.Extension.ToLower() == ".doc" || f.Extension.ToLower() == ".docx"); }).Count() > 10)
                        {
                            ShowLog(fs.Name + "文件夹下有多个文件,可能不是同一作者,请打开分别整理", 3);
                            continue;
                        }
                        //存储该文件夹内所有文件生成的pdf文件的路径,最后要合并
                        List<string> pdfNames = new List<string>();
                        foreach (FileInfo f in fileList)
                        {
                            string[] arrayStr = FileOprate_Chunpan(f, out backpdf);
                            string afterCopyFilename = CopyFile(f, desDirName);
                            string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}','{4}',{5},\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, afterCopyFilename, f.Extension, arrayStr[1], arrayStr[0], 0, f.FullName, 0, "null", desDirName + "\\" + afterCopyFilename);
                            dbHelper.ExecuteNonQuery(sql, null);
                            if (arrayStr[0] != "是" || arrayStr[1] != "是")
                            {
                                ViewModel.CurFail++;
                                ViewModel.TotalFail++;
                            }
                            else
                            {
                                ViewModel.CurSuc++;
                                ViewModel.TotalSuc++;
                            }
                            ViewModel.UnTidy--;


                            pdfNames.Add(backpdf);
                            XiaoYangSum++;
                        }
                        //开始合并该任务内所有pdf文件
                        string outPdfName;
                        if (pdfNames.Contains(""))//文件整理异常
                            outPdfName = "";
                        else if (pdfNames.All(a => { return a == null; }))//没有一个文件提取出有效信息,则不必合并pdf了
                        {
                            if (!File.Exists(TEMP_PDF))
                                Voidpdf();
                            outPdfName = TEMP_PDF;
                        }
                        else//提取出了有效信息pdf
                        {
                            outPdfName = ViewModel.Upload_Path + "\\" + ViewModel.Begin_Code + "[整].pdf";
                            CombineMultiplePDFs(pdfNames, ref outPdfName);
                        }

                        if (outPdfName != "")//合并成功
                        {
                            fs.HasTidy = 1;
                            backpdf = outPdfName;
                        }
                        else//合并失败
                        {
                            fs.HasTidy = -1;
                            backpdf = "";
                        }
                    }
                    #endregion

                    #region 勾选项是文件
                    else
                    {
                        FileInfo file = new FileInfo(fs.FullPath);
                        string[] arrayStr = FileOprate_Chunpan(file, out backpdf);
                        string afterCopyFilename = CopyFile(file, desDirName);
                        //文件信息入库
                        string sql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}','{4}',{5},\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, afterCopyFilename, file.Extension, arrayStr[1], arrayStr[0], 0, file.FullName, 0, "null", desDirName + "\\" + afterCopyFilename);
                        dbHelper.ExecuteNonQuery(sql, null);
                        if (arrayStr[0] != "是" || arrayStr[1] != "是")
                        {
                            ViewModel.CurFail++;
                            ViewModel.TotalFail++;
                            fs.HasTidy = -1;
                        }
                        else
                        {
                            ViewModel.CurSuc++;
                            ViewModel.TotalSuc++;
                            fs.HasTidy = 1;
                        }
                        if (backpdf == null)//如果没有提取出有效信息
                        {
                            if (!File.Exists(TEMP_PDF))
                                Voidpdf();
                            backpdf = TEMP_PDF;
                        }
                        ViewModel.UnTidy--;
                        XiaoYangSum++;

                    }
                    #endregion

                    int TotalPage = 0;
                    if (backpdf != "")
                    {
                        WaterMarkPdf(ref backpdf, out TotalPage);
                        ShowLog(ViewModel.Begin_Code + "生成任务成功", 1);
                    }
                    else
                        ShowLog(ViewModel.Begin_Code + "生成任务失败", 3);

                    //整理出的任务信息入库
                    string str = string.Format("insert into XW_FileOrderinfo(编号,路径,提交否,文件名,年度,级别,版权反馈否,保密否,是否签名,是否授权,删除字样,提取页数,小样数) values('{0}','{1}','否','{2}','{3}','硕士','是','否','否','否','否',{4},{5})", ViewModel.Begin_Code, desDirName, backpdf, year, TotalPage, XiaoYangSum);
                    dbHelper.ExecuteNonQuery(str, null);

                    fs.Checked = false;
                    //更新配置文件编号
                    var num = Convert.ToInt64(ViewModel.Begin_Code);
                    ViewModel.Begin_Code = (num + 1).ToString();
                    ConfigHelper.SetValue("BEGIN_CODE", ViewModel.Begin_Code);
                }

                ShowLog("生成任务完毕", 2);
            }
            catch (Exception ee)
            {
                ShowLog(ee.Message, 3);
                Utility.Log.TextLog.WritwLog(ee.Message);
            }
        }
        /// <summary>
        /// 刊盘单独生成
        /// </summary>
        public void KanPan_Alone()
        {
            try
            {
                ShowLog("开始生成任务", 2);
                var checkedItem = ViewModel.Models.Where(m => { return m.Checked == true && m.Name != ".."; });
                if (checkedItem.Count() == 0)
                {
                    ShowLog("没有选择任何文件", 3);
                    return;
                }
                int readWordNum = Convert.ToInt32(ConfigHelper.GetValue("Word_Summary_Num"));
                int readPdfNum = Convert.ToInt32(ConfigHelper.GetValue("PDF_Summary_PageNum"));

                Reader reader = new Reader(readWordNum, readPdfNum);
                List<string> listPdf = new List<string>();

                string desDirName = ViewModel.AfterTydyPath + "\\" + ViewModel.Begin_Code;
                if (!Directory.Exists(desDirName))
                    Directory.CreateDirectory(desDirName);
                int XiaoYangSum = 0;
                foreach (Model_FileSystem fs in checkedItem)
                {
                    #region 整理文件夹
                    if (fs.Type == SystemType.Dir)
                    {
                        foreach (FileInfo file in Traverse(fs.FullPath))
                        {
                            string backpdf;
                            //整理刊盘文件
                            string[] backStr = FileOprate_KanPan(file, reader, out backpdf);
                            string afterCopyFilename = CopyFile(file, desDirName);
                            //文件信息入库
                            string insertSql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, afterCopyFilename, backStr[0], backStr[1], 0, file.FullName, backStr[2].Replace('"', ' '), 0, "null", desDirName + "\\" + afterCopyFilename);
                            try
                            {
                                dbHelper.ExecuteNonQuery(insertSql, null);
                            }
                            catch (Exception e)
                            {
                                backpdf = "";
                                insertSql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, file.Name, backStr[0], '否', 0, file.FullName, "", 0, "null", desDirName + "\\" + file.Name);
                                dbHelper.ExecuteNonQuery(insertSql, null);
                                ShowLog("本地数据库写入失败，文件可能包含乱码", 3);
                            }
                            listPdf.Add(backpdf);
                            if (backStr[0] != "是" || backStr[1] != "是")
                            {
                                ViewModel.CurFail++;
                                ViewModel.TotalFail++;
                            }
                            else
                            {
                                ViewModel.CurSuc++;
                                ViewModel.TotalSuc++;
                            }
                            ViewModel.UnTidy--;
                            XiaoYangSum++;
                        }
                    }
                    #endregion

                    #region 整理单个文件
                    else
                    {
                        FileInfo file = new FileInfo(fs.FullPath);
                        string backpdf;
                        string[] backStr = FileOprate_KanPan(file, reader, out backpdf);
                        string afterCopyFilename = CopyFile(file, desDirName);
                        string insertSql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, afterCopyFilename, backStr[0], backStr[1], 0, file.FullName, backStr[2].Replace('"', ' '), 0, "null", desDirName + "\\" + afterCopyFilename);
                        try
                        {
                            dbHelper.ExecuteNonQuery(insertSql, null);
                        }
                        catch (Exception e)
                        {
                            backpdf = "";
                            insertSql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, file.Name, backStr[0], '否', 0, file.FullName, "", 0, "null", desDirName + "\\" + file.Name);
                            dbHelper.ExecuteNonQuery(insertSql, null);
                            ShowLog("本地数据库写入失败，文件可能包含乱码", 3);
                        }
                        listPdf.Add(backpdf);

                        //整理成功更新工具界面
                        if (backStr[0] != "是" || backStr[1] != "是")
                        {
                            fs.HasTidy = -1;
                            ViewModel.TotalFail++;
                            ViewModel.CurFail++;
                        }
                        //整理失败更新工具界面
                        else
                        {
                            fs.HasTidy = 1;
                            ViewModel.CurSuc++;
                            ViewModel.TotalSuc++;
                        }
                        ViewModel.UnTidy--;
                        XiaoYangSum++;
                    }
                    #endregion
                    fs.Checked = false;
                }

                string finalPdf;
                if (listPdf.Contains(""))
                    finalPdf = "";
                else if (listPdf.All(a => { return a == null; }))//没有一个文件提取出有效信息,则不必合并pdf了
                {
                    if (!File.Exists(TEMP_PDF))
                        Voidpdf();
                    finalPdf = TEMP_PDF;
                }
                else
                {
                    //合并整理好的pdf
                    finalPdf = ViewModel.Upload_Path + "\\" + ViewModel.Begin_Code + "[整].pdf";
                    CombineMultiplePDFs(listPdf, ref finalPdf);
                }
                int TotalPage = 0;
                //pdf加水印
                if (finalPdf != "")
                    WaterMarkPdf(ref finalPdf, out TotalPage);
                string sql = "insert into XW_FileOrderinfo(编号,文件名,删除字样,保密否,提取页数,小样数,路径) values('{0}','{1}','{2}','{3}',{4},{5},'{6}')";
                sql = string.Format(sql, ViewModel.Begin_Code, finalPdf, "否", "否", TotalPage, XiaoYangSum, desDirName);
                dbHelper.ExecuteNonQuery(sql, null);
                //插入当前任务状态
                sql = "insert into db_State values('" + ViewModel.Begin_Code + "','否')";
                dbHelper.ExecuteNonQuery(sql, null);

                ShowLog(ViewModel.Begin_Code + "生成任务成功", 1);
                //计算下一个任务的流水号
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
                ConfigHelper.SetValue("BEGIN_CODE_Kanpan_" + ViewModel.GongHao, ViewModel.Begin_Code);

            }
            catch (Exception ee)
            {
                ShowLog(ee.Message, 3);
                Utility.Log.TextLog.WritwLog(ee.Message);
            }
        }

        /// <summary>
        /// 刊盘批量生成
        /// </summary>
        public void KanPan_Many()
        {
            try
            {
                ShowLog("开始批量生成", 2);
                var checkedItem = ViewModel.Models.Where(m => { return m.Checked == true && m.Name != ".."; });
                if (checkedItem.Count() == 0)
                {
                    ShowLog("没有选择任何文件", 3);
                    return;
                }
                int readWordNum = Convert.ToInt32(ConfigHelper.GetValue("Word_Summary_Num"));
                int readPdfNum = Convert.ToInt32(ConfigHelper.GetValue("PDF_Summary_PageNum"));
                Reader reader = new Reader(readWordNum, readPdfNum);
                string insertSql;
                int year = DateTime.Now.Year;


                foreach (Model_FileSystem fs in checkedItem)
                {
                    if (tokenSource.Token.IsCancellationRequested)
                    {
                        ShowLog("整理停止", 2);
                        return;
                    }
                    string desDirName = ViewModel.AfterTydyPath + "\\" + ViewModel.Begin_Code;
                    if (!Directory.Exists(desDirName))
                        Directory.CreateDirectory(desDirName);
                    int XiaoYangSum = 0;
                    #region 包含一个文件夹的任务

                    string backPdf;
                    if (fs.Type == SystemType.Dir)
                    {
                        var fileList = Traverse(fs.FullPath);

                        if (fileList.Where(f => { return (f.Extension.ToLower() == ".pdf" || f.Extension.ToLower() == ".doc" || f.Extension.ToLower() == ".docx"); }).Count() > 10)
                        {
                            ShowLog(fs.Name + "文件夹下有多个文件,可能不是同一作者,请打开分别整理", 3);
                            continue;
                        }
                        Utility.Log.TextLog.WritwLog("111");
                        List<string> listPdf = new List<string>();
                        //foreach (FileInfo file in new DirectoryInfo(fs.FullPath).GetFiles())
                        foreach (FileInfo file in fileList)
                        {
                            string outpdf;
                            string[] backStr = FileOprate_KanPan(file, reader, out outpdf);
                            string afterCopyFilename = CopyFile(file, desDirName);
                            insertSql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, afterCopyFilename, backStr[0], backStr[1], 0, file.FullName, backStr[2].Replace('"', ' '), 0, "null", desDirName + "\\" + afterCopyFilename);
                            try
                            {
                                dbHelper.ExecuteNonQuery(insertSql, null);
                            }
                            catch (Exception e)
                            {
                                outpdf = "";
                                insertSql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, file.Name, backStr[0], '否', 0, file.FullName, "", 0, "null", desDirName + "\\" + file.Name);
                                dbHelper.ExecuteNonQuery(insertSql, null);
                                ShowLog("本地数据库写入失败，文件可能包含乱码", 3);
                            }
                            if (backStr[0] != "是" || backStr[1] != "是")
                            {
                                ViewModel.CurFail++;
                                ViewModel.TotalFail++;
                            }
                            else
                            {
                                ViewModel.CurSuc++;
                                ViewModel.TotalSuc++;
                            }
                            ViewModel.UnTidy--;
                            listPdf.Add(outpdf);
                            XiaoYangSum++;
                        }
                        if (listPdf.Contains(""))
                        {
                            backPdf = "";
                            fs.HasTidy = -1;
                        }
                        else if (listPdf.All(a => { return a == null; }))//没有一个文件提取出有效信息,则不必合并pdf了
                        {
                            if (!File.Exists(TEMP_PDF))
                                Voidpdf();
                            backPdf = TEMP_PDF;
                        }
                        else
                        {
                            backPdf = ViewModel.Upload_Path + "\\" + ViewModel.Begin_Code + "[整].pdf";
                            CombineMultiplePDFs(listPdf, ref backPdf);
                            fs.HasTidy = 1;
                        }
                    }
                    #endregion

                    #region 整理单个文件
                    else
                    {
                        FileInfo file = new FileInfo(fs.FullPath);
                        //整理刊盘文件
                        string[] backStr = FileOprate_KanPan(file, reader, out backPdf);
                        string afterCopyFilename = CopyFile(file, desDirName);
                        insertSql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, afterCopyFilename, backStr[0], backStr[1], 0, file.FullName, backStr[2].Replace('"', ' '), 0, "null", desDirName + "\\" + afterCopyFilename);
                        insertSql.Replace("\0", "");
                        Utility.Log.TextLog.WritwLog("刊盘批量整理sql：" + insertSql);
                        try
                        {
                            dbHelper.ExecuteNonQuery(insertSql, null);
                        }
                        catch (Exception e)
                        {
                            insertSql = string.Format("insert into db_File values('{0}',\"{1}\",'{2}','{3}',{4},\"{5}\",\"{6}\",{7},{8},\"{9}\")", ViewModel.Begin_Code, file.Name, backStr[0], '否', 0, file.FullName, "", 0, "null", desDirName + "\\" + file.Name);
                            dbHelper.ExecuteNonQuery(insertSql, null);
                            ShowLog("本地数据库写入失败，文件可能包含乱码", 3);
                        }


                        if (backStr[0] != "是" || backStr[1] != "是")
                        {
                            ViewModel.CurFail++;
                            ViewModel.TotalFail++;
                            fs.HasTidy = -1;
                        }
                        else
                        {
                            fs.HasTidy = 1;
                            ViewModel.CurSuc++;
                            ViewModel.TotalSuc++;
                        }
                        if (backPdf == null)//如果没有提取出有效信息
                        {
                            if (!File.Exists(TEMP_PDF))
                                Voidpdf();
                            backPdf = TEMP_PDF;
                        }
                        ViewModel.UnTidy--;
                        XiaoYangSum++;
                    }
                    #endregion

                    int TotalPage = 0;
                    if (backPdf != "")
                    {
                        WaterMarkPdf(ref backPdf, out TotalPage);
                        ShowLog(ViewModel.Begin_Code + "生成任务成功!", 1);
                    }
                    else
                        ShowLog(ViewModel.Begin_Code + "生成任务失败!", 3);
                    insertSql = "insert into XW_FileOrderinfo(编号,文件名,删除字样,保密否,提取页数,小样数,路径) values('{0}','{1}','{2}','{3}',{4},{5},'{6}')";
                    insertSql = string.Format(insertSql, ViewModel.Begin_Code, backPdf, "否", "否", TotalPage, XiaoYangSum, desDirName);
                    dbHelper.ExecuteNonQuery(insertSql, null);
                    //插入当前任务状态
                    insertSql = "insert into db_State values('" + ViewModel.Begin_Code + "','否')";
                    dbHelper.ExecuteNonQuery(insertSql, null);

                    //生成下个刊盘编号,更新配置文件
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
                    ConfigHelper.SetValue("BEGIN_CODE_Kanpan_" + ViewModel.GongHao, ViewModel.Begin_Code);
                    fs.Checked = false;

                }

                ShowLog("批量生成结束", 2);
            }
            catch (Exception ee)
            {
                Utility.Log.TextLog.WritwLog(ee.Message);
                ShowLog("整理出错:" + ee.Message, 3);
            }
        }

        /// <summary>
        /// 重整刊盘错误
        /// </summary>
        public void KanPan_Error()
        {
            try
            {
                ShowLog("开始重新整理失败刊盘任务", 2);
                var checkedItem = ViewModel.Models.Where(m => m.Checked == true);
                int readWordNum = Convert.ToInt32(ConfigHelper.GetValue("Word_Summary_Num"));
                int readPdfNum = Convert.ToInt32(ConfigHelper.GetValue("PDF_Summary_PageNum"));
                Reader reader = new Reader(readWordNum, readPdfNum);
                //for(int i=0;i<checkedItem.Count();i++)
                foreach (Model_FileSystem fs in checkedItem)
                {
                    System.Data.DataTable dt = dbHelper.ExecuteDataTable("select 路径 from db_File where 编号='" + fs.Name + "'", null);
                    //bool haha = true;//如果这批任务中有一个文件没有整理成功,则该任务就不算整理成功
                    List<string> listPdf = new List<string>();
                    int XiaoYangSum = dt.Rows.Count;
                    foreach (System.Data.DataRow row in dt.Rows)
                    {
                        FileInfo file = new FileInfo(row[0].ToString());
                        string outpdf;
                        string[] backStr = FileOprate_KanPan(file, reader, out outpdf);
                        string updateSql = string.Format("update db_File set 提取='{0}',可读='{1}',摘要=\"{2}\",起始页={3},结束页={4} where 路径=\"{5}\"", backStr[0], backStr[1], backStr[2], 0, 0, row[0].ToString());
                        dbHelper.ExecuteNonQuery(updateSql, null);
                        listPdf.Add(outpdf);
                    }
                    string finalPdf;
                    if (listPdf.Contains(""))
                    {
                        ShowLog("任务编号" + fs.Name + "重新整理失败!", 3);
                        continue;
                    }
                    else if (listPdf.All(a => { return a == null; }))//没有一个文件提取出有效信息,则不必合并pdf了
                    {
                        if (!File.Exists(TEMP_PDF))
                            Voidpdf();
                        finalPdf = TEMP_PDF;
                    }
                    else
                    {
                        finalPdf = ViewModel.Upload_Path + "\\" + fs.Name + "[整].pdf";
                        CombineMultiplePDFs(listPdf, ref finalPdf);
                    }
                    int TotalPage = 0;
                    WaterMarkPdf(ref finalPdf, out TotalPage, fs.Name);
                    string sql = "update XW_FileOrderinfo set 小样数=" + XiaoYangSum + ",提取页数=" + TotalPage + ", 文件名='" + finalPdf + "' where 编号='" + fs.Name + "'";
                    dbHelper.ExecuteNonQuery(sql, null);
                    RemoveItem(fs);
                    ShowLog("任务编号" + fs.Name + "重新整理成功!", 2);

                }

            }
            catch (Exception ee)
            {
                Utility.Log.TextLog.WritwLog(ee.Message);
            }
        }
        /// <summary>
        /// 重整纯盘错误
        /// </summary>
        public void ChunPan_Error()
        {
            try
            {
                ShowLog("开始重新整理失败纯盘任务", 2);
                var checkedItem = ViewModel.Models.Where(m => m.Checked == true);
                foreach (Model_FileSystem fs in checkedItem)
                {
                    System.Data.DataTable dt = dbHelper.ExecuteDataTable("select 路径 from db_File where 编号='" + fs.Name + "'", null);
                    //bool haha = true;//如果这批任务中有一个文件没有整理成功,则该任务就不算整理成功
                    List<string> listPdf = new List<string>();
                    int XiaoYangSum = dt.Rows.Count;
                    foreach (System.Data.DataRow row in dt.Rows)
                    {
                        FileInfo file = new FileInfo(row[0].ToString());
                        string outpdf;
                        string[] backStr = FileOprate_Chunpan(file, out outpdf);
                        string updateSql = string.Format("update db_File set 提取='{0}',可读='{1}',起始页={2},结束页={3} where 路径=\"{4}\"", backStr[0], backStr[1], 0, 0, row[0].ToString());
                        dbHelper.ExecuteNonQuery(updateSql, null);
                        listPdf.Add(outpdf);

                    }
                    string finalPdf;
                    if (listPdf.Contains(""))
                    {
                        ShowLog("任务编号" + fs.Name + "重新整理失败!", 3);
                        continue;
                    }
                    else if (listPdf.All(a => { return a == null; }))//没有一个文件提取出有效信息,则不必合并pdf了
                    {
                        if (!File.Exists(TEMP_PDF))
                            Voidpdf();
                        finalPdf = TEMP_PDF;
                    }
                    else
                    {
                        finalPdf = ViewModel.Upload_Path + "\\" + fs.Name + "[整].pdf";
                        CombineMultiplePDFs(listPdf, ref finalPdf);
                    }
                    int TotalPage = 0;
                    WaterMarkPdf(ref finalPdf, out TotalPage, fs.Name);
                    string sql = "update XW_FileOrderinfo set 小样数=" + XiaoYangSum + ", 提取页数=" + TotalPage + ", 文件名='" + finalPdf + "' where 编号='" + fs.Name + "'";
                    dbHelper.ExecuteNonQuery(sql, null);
                    RemoveItem(fs);
                    ShowLog("任务编号" + fs.Name + "重新整理成功!", 2);

                }
            }
            catch (Exception ee)
            {
                Utility.Log.TextLog.WritwLog(ee.Message);
            }
        }

        public List<FileInfo> Traverse(string sPathName)
        {
            List<FileInfo> list = new List<FileInfo>();
            //创建一个队列用于保存子目录
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
                    if (fi.Name.StartsWith("~$"))
                        continue;
                    list.Add(fi);
                }
            }
            return list;
        }


        /// <summary>
        /// 处理纯盘任务文件
        /// </summary>
        /// <param name="file">要处理的文件</param>
        /// <param name="page"></param>
        /// <returns>string[]{提取,可读}</returns>
        private string[] FileOprate_Chunpan(FileInfo file, out string BackPdfPath)
        {

            string pdfPath = "";
            string[] result;
            if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".wps" || file.Extension.ToLower() == ".rtf" || file.Extension.ToLower() == ".dot")
            {
                Task<string> word2pdf;
                if(CUSTOM_DEFINE)
                    word2pdf = new Task<string>(MadeDefinedWord, new DefinedWordPra { path = file.FullName,front=WORD_FRONT_NUM,back=WORD_BACK_NUM });
                else
                    word2pdf = new Task<string>(MadeWord, file.FullName);
                Utility.Log.TextLog.WritwLog("end MadeWord");
                word2pdf.Start();
                word2pdf.Wait(360 * 1000);
                Utility.Log.TextLog.WritwLog("FileOprate_Chunpan1");
                if (word2pdf.IsCompleted)
                {
                    pdfPath = word2pdf.Result;
                    Utility.Log.TextLog.WritwLog("FileOprate_Chunpan2");
                    if (pdfPath != "")
                    {
                        result = new string[2] { "是", "是" };
                    }
                    else
                    {
                        result = new string[2] { "否", "否" };
                    }
                }
                else
                {
                    result = new string[2] { "否", "否" };
                    Reader.KillWord();
                }
            }
            else if (file.Extension.ToLower() == ".pdf")
            {
                if (CUSTOM_DEFINE)
                    pdfPath = MadeDefinedPdf(file.FullName, PDF_FRONT_NUM, PDF_BACK_NUM);
                else
                    pdfPath = MadePdf(file.FullName);
                if (pdfPath != "")
                {
                    result = new string[2] { "是", "是" };
                }
                else
                {
                    result = new string[2] { "否", "否" };
                }
            }
            else if (file.Extension.ToLower() == ".jpg" || file.Extension.ToLower() == ".jpeg" || file.Extension.ToLower() == ".tif" || file.Extension.ToLower() == ".bmp" || file.Extension.ToLower() == ".png")
            {
                pdfPath = MadeJpg(file.FullName);
                if (pdfPath != "")
                {
                    result = new string[2] { "是", "是" };
                }
                else
                {
                    result = new string[2] { "否", "否" };
                }
            }
            else
            {
                result = new string[2] { "不需", "不需" };
                pdfPath = "";
            }
            BackPdfPath = pdfPath;
            Utility.Log.TextLog.WritwLog("FileOprate_Chunpan3");
            return result;
        }
        /// <summary>
        /// 刊盘 文件处理
        /// </summary>
        /// <param name="file">要整理的文件</param>
        /// <param name="reader">读取word或pdf的方法</param>
        /// <returns>在pdf中的页数, 和string[]{提取,可读,摘要}</returns>
        private string[] FileOprate_KanPan(FileInfo file, Reader reader, out string BackPdfPath)
        {

            string text = "";

            string[] result = new string[3];

            string pdfPath = null;
            //如果文件是WORD
            if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".wps" || file.Extension.ToLower() == ".rtf" || file.Extension.ToLower() == ".dot")
            {
                if (file.Name.StartsWith("~$"))
                {
                    BackPdfPath = "不需";
                    return new string[3] { "不需", "不需", "" };
                }
                //执行提取word文件的后台任务
                Task<string> word2pdf;
                if(CUSTOM_DEFINE)
                    word2pdf = new Task<string>(MadeDefinedWord, new DefinedWordPra { path = file.FullName,front=WORD_FRONT_NUM,back=WORD_BACK_NUM });
                else
                    word2pdf = new Task<string>(MadeWord, file.FullName);
                word2pdf.Start();
                //等待60秒,如果等待时间内还没完成,则算整理失败
                word2pdf.Wait(360 * 1000);
                if (word2pdf.IsCompleted)
                {
                    pdfPath = word2pdf.Result;

                    if (pdfPath != "")
                    {
                        result[0] = "是";
                    }
                    else
                    {
                        result[0] = "否";
                    }
                }
                else
                {
                    result[0] = "否";
                    pdfPath = "";
                    Reader.KillWord();
                }
                reader.ReadHandler = reader.ReadWord;
            }
            //整理pdf文件
            else if (file.Extension.ToLower() == ".pdf")
            {
                if (CUSTOM_DEFINE)
                    pdfPath = MadeDefinedPdf(file.FullName, PDF_FRONT_NUM, PDF_BACK_NUM);
                else
                    pdfPath = MadePdf(file.FullName);
                if (pdfPath != "")
                    result[0] = "是";
                else
                    result[0] = "否";
                reader.ReadHandler = reader.ReadPdf;
            }
            //整理图片
            else if (file.Extension.ToLower() == ".jpg" || file.Extension.ToLower() == ".jpeg" || file.Extension.ToLower() == ".tif" || file.Extension.ToLower() == ".bmp" || file.Extension.ToLower() == ".png")
            {
                pdfPath = MadeJpg(file.FullName);
                if (pdfPath != "")
                    result[0] = "是";
                else
                    result[0] = "否";
            }
            else
            {
                BackPdfPath = "";
                return new string[3] { "不需", "不需", "" };
            }

            try
            {
                //读取word文件和pdf文件的文字摘要
                if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".wps" || file.Extension.ToLower() == ".rtf" || file.Extension.ToLower() == ".pdf")
                {
                    text = reader.ReadWithTimeout(file.FullName);
                }
            }
            catch (Exception e1)
            {
                result[1] = "否";
            }
            if (text == "文件读取异常" || text == "文件读取超时" || text == "文件内容为乱码" || text == "")//如果是文件读取超时需要注意杀死pdftotext.exe，否则文件占用无法移动
            {
                //ShowLog(file.Name + text, 3);
                result[1] = "否";
                result[2] = file.Name;

            }

            else
            {
                result[1] = "是";
                result[2] = file.Name + text;
            }
            BackPdfPath = pdfPath;
            Utility.Log.TextLog.WritwLog("BackPdfPath:" + BackPdfPath);

            return result;
        }


        /// <summary>
        /// 整理图片为pdf
        /// </summary>
        /// <param name="path">图片路径</param>
        /// <returns>生成的pdf路径</returns>
        private string MadeJpg(string jpgfile)
        {
            //整理图片文件,就是把图片转成pdf
            try
            {
                string jpgname = System.IO.Path.GetFileName(jpgfile);
                string pdf = ViewModel.Upload_Path + "\\" + jpgname.Substring(0, jpgname.LastIndexOf('.')) + "[整].pdf";
                var document = new Document(iTextSharp.text.PageSize.A4, 25, 25, 25, 25);
                using (var stream = new FileStream(pdf, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    PdfWriter.GetInstance(document, stream);
                    document.Open();
                    using (var imageStream = new FileStream(jpgfile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        var image = iTextSharp.text.Image.GetInstance(imageStream);
                        if (image.Height > iTextSharp.text.PageSize.A4.Height - 25)
                        {
                            image.ScaleToFit(iTextSharp.text.PageSize.A4.Width - 25, iTextSharp.text.PageSize.A4.Height - 25);
                        }
                        else if (image.Width > iTextSharp.text.PageSize.A4.Width - 25)
                        {
                            image.ScaleToFit(iTextSharp.text.PageSize.A4.Width - 25, iTextSharp.text.PageSize.A4.Height - 25);
                        }
                        image.Alignment = iTextSharp.text.Image.ALIGN_MIDDLE;
                        document.Add(image);
                    }
                    document.Close();
                }
                return pdf;
            }
            catch (Exception e)
            {
                Utility.Log.TextLog.WritwLog(e.Message + "  文件名是:" + jpgfile, true);
                return "";
            }
        }

        #region 提取word非正文页1.0
        /// <summary>
        /// 从word中提取pdf文件,不加后五页
        /// </summary>
        /// <param name="wordFilePath"></param>
        /// <returns></returns>
        //public string MadeWord(object wordFilePath)
        //{
        //    Word.Application _app = null;
        //    Word.Document document = null;
        //    string filename = System.IO.Path.GetFileName(wordFilePath.ToString());
        //    string fullPdfName = ViewModel.Upload_Path + "\\" + filename.Substring(0, filename.LastIndexOf('.')) + "[整].pdf";
        //    try
        //    {

        //        _app = new Word.Application();
        //        Word.Documents d = _app.Documents;
        //        document = d.Open(wordFilePath, false, false, false, ref G_missing, G_missing, false, G_missing, G_missing, G_missing, G_missing, false, false, G_missing, true, G_missing);
        //        Word.Document P_document = d.Add(ref G_missing, G_missing, ref G_missing);

        //        Utility.Log.TextLog.WritwLog("打开word文件成功");
        //        var tocs = document.TablesOfContents;
        //        Utility.Log.TextLog.WritwLog("获取word目录");
        //        //如果此word文件没有目录
        //        if (tocs.Count == 0)
        //        {
        //            Utility.Log.TextLog.WritwLog("word目录是零");
        //            int to = Convert.ToInt32(ConfigHelper.GetValue("Word_NoCatalog_ExNum"));
        //            document.ExportAsFixedFormat(fullPdfName, Word.WdExportFormat.wdExportFormatPDF, false, Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Word.WdExportRange.wdExportFromTo, 1, to);
        //            //document.Close();
        //            //_app.Quit();
        //            //return fullPdfName;

        //        }
        //        else
        //        {
        //            var toc = tocs[1];
        //            Utility.Log.TextLog.WritwLog("找到word目录");
        //            Word.Range ran = toc.Range.Previous(Word.WdUnits.wdSection);
        //            if (ran == null)
        //            {
        //                int to = Convert.ToInt32(ConfigHelper.GetValue("Word_NoCatalog_ExNum"));
        //                document.ExportAsFixedFormat(fullPdfName, Word.WdExportFormat.wdExportFormatPDF, false, Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Word.WdExportRange.wdExportFromTo, 1, to);
        //                //document.Close();
        //                //_app.Quit();
        //                //return fullPdfName;
        //            }
        //            else
        //            {
        //                while (ran != null)
        //                {
        //                    ran.Copy();
        //                    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref G_missing, ref G_missing);
        //                    try
        //                    {
        //                        P_document.ActiveWindow.Selection.Paste();
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        string text = P_document.Tables[1].Range.Text;
        //                        P_document.Tables[1].Delete();
        //                        P_document.Range(0, 0).Text = text;
        //                        P_document.ActiveWindow.Selection.Paste();
        //                    }
        //                    ran = ran.Previous(Word.WdUnits.wdSection);
        //                }
        //                P_document.ExportAsFixedFormat(fullPdfName, Word.WdExportFormat.wdExportFormatPDF);//word前几页转成pdf文件
        //            }
        //            //object saveOption= Word.WdSaveOptions.wdDoNotSaveChanges;


        //        }

        //        (document as Word._Document).Close(ref saveOption, ref G_missing, ref G_missing);
        //        Marshal.FinalReleaseComObject(document);
        //        (P_document as Word._Document).Close(ref saveOption, ref G_missing, ref G_missing);
        //        Marshal.FinalReleaseComObject(P_document);
        //        Marshal.FinalReleaseComObject(d);
        //        (_app as Word._Application).Quit(ref G_missing, ref G_missing, ref G_missing);
        //        Marshal.FinalReleaseComObject(_app);

        //        return fullPdfName;
        //    }
        //    catch (Exception e)
        //    {
        //        //if (document != null)
        //        //    (document as Word._Document).Close(ref saveOption, ref G_missing, ref G_missing);
        //        //if (_app != null)
        //        //    (_app as Word._Application).Quit(ref G_missing, ref G_missing, ref G_missing);
        //        Utility.Log.TextLog.WritwLog(e.Message);
        //        ShowLog(e.Message, 3);
        //        return "";
        //    }
        //}
        #endregion

        #region 提取word非正文页2.0
        /// <summary>
        /// 提取前5页除摘要、致谢以外的内容,提取后5页中的图片和摘要内容
        /// </summary>
        /// <param name="wordFilePath"></param>
        /// <returns>"":异常 null:没有提取出有效信息</returns>
        //public string MadeWord(object wordFilePath)
        //{
        //    string filename = System.IO.Path.GetFileName(wordFilePath.ToString());
        //    Word.Application _app = null;
        //    Word.Document document = null;
        //    object missing = System.Reflection.Missing.Value;
        //    try
        //    {
        //        string fullPdfName = ViewModel.Upload_Path + "\\" + filename.Substring(0, filename.LastIndexOf('.')) + "[整].pdf";
        //        _app = new Word.Application();
        //        document = _app.Documents.Open(wordFilePath.ToString(), false, false, false, ref missing, missing, false, missing, missing, missing, missing, false, false, missing, true, missing);

        //        Word.Document P_document = _app.Documents.Add(ref missing, ref missing, ref missing);
        //        object What = Word.WdGoToItem.wdGoToSection;
        //        object Which = Word.WdGoToDirection.wdGoToLast;

        //        #region 处理前5页
        //        foreach (Word.Section section in document.Sections)
        //        {
        //            var range = section.Range;
        //            if (range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) > WORD_FRONT_NUM)
        //                break;

        //            bool needCopy = true;
        //            for (int j = 1; j <= range.Sentences.Count; j++)
        //            {
        //                string text = range.Sentences[j].Text.Trim().Replace(" ", "");
        //                if (Front5Array.Contains(text))
        //                {
        //                    needCopy = false;
        //                    break;
        //                }
        //            }
        //            if (needCopy)
        //            {
        //                range.Copy();
        //                P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                P_document.ActiveWindow.Selection.Paste();
        //            }

        //        }
        //        #endregion

        //        #region 处理后5页
        //        int pages = document.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages);//总页数
        //        if (pages > WORD_FRONT_NUM)
        //        {
        //            int secCount = document.Sections.Count;
        //            int xxCount = pages < WORD_FRONT_NUM + WORD_BACK_NUM ? pages - WORD_FRONT_NUM : WORD_BACK_NUM;
        //            for (int i = secCount; i > 0; i--)
        //            {
        //                var range = document.Sections[i].Range;
        //                if (range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) <= pages - xxCount)
        //                    break;
        //                string text = range.Text.Trim().Replace(" ", "");
        //                if (Back5Array.Any(a => { return text.Contains(a); }))
        //                {
        //                    range.Copy();
        //                    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                    P_document.ActiveWindow.Selection.Paste();
        //                }
        //                else
        //                {
        //                    foreach (Word.InlineShape inlShape in range.InlineShapes)
        //                    {
        //                        if (inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapePicture) || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
        //                            || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject))
        //                        {
        //                            inlShape.Range.Copy();
        //                            P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                            P_document.ActiveWindow.Selection.Paste();
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        #endregion

        //        if (P_document.Words.Count > 1)//判断是否提取出了有效信息,如果没有则返回null
        //            P_document.ExportAsFixedFormat(fullPdfName, Word.WdExportFormat.wdExportFormatPDF);
        //        else
        //            fullPdfName = null;
        //        P_document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        //        document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        //        _app.Quit();
        //        return fullPdfName;
        //    }
        //    catch (Exception e)
        //    {
        //        ShowLog("MadeWord异常:" + e.Message,3);
        //        Utility.Log.TextLog.WritwLog("整理" + filename + "失败:" + e.Message, true);
        //        return "";
        //    }
        //}
        #endregion

        //#region 提取word非正文页3.0
        //string[] Front5WordFetch = ConfigHelper.GetValue("Front5DocFetchWord").Split(new char[] { ',' });
        //public string MadeWord(object wordFilePath)
        //{
        //    Utility.Log.TextLog.WritwLog("madeword1");
        //    string filename = System.IO.Path.GetFileName(wordFilePath.ToString());
        //    Utility.Log.TextLog.WritwLog("madeword2");
        //    Word.Application _app = null;
        //    Word.Document document = null;
        //    object missing = System.Reflection.Missing.Value;
        //    try
        //    {
        //        //整理后生成的文件名
        //        Utility.Log.TextLog.WritwLog("madeword3");
        //        string fullPdfName = ViewModel.Upload_Path + "\\" + filename.Substring(0, filename.LastIndexOf('.')) + "[整].pdf";
        //        Utility.Log.TextLog.WritwLog("madeword3-1" + wordFilePath);
        //        _app = new Word.Application();
        //        Utility.Log.TextLog.WritwLog("madeword3-2");
        //        document = _app.Documents.Open(wordFilePath.ToString(), false, false, false, ref missing, missing, false, missing, missing, missing, missing, false, false, missing, true, missing);
        //        Utility.Log.TextLog.WritwLog("madeword3-3");
        //        Word.Document P_document = _app.Documents.Add(ref missing, ref missing, ref missing);
        //        object What = Word.WdGoToItem.wdGoToSection;
        //        object Which = Word.WdGoToDirection.wdGoToLast;
        //        #region 在第WORD_FRONT_NUM页末尾插入分节符,在第WORD_BACK_NUM页前插入分节符,避免提取过多页数
        //        object Name = (object)(WORD_FRONT_NUM + 1);
        //        object What1 = Word.WdGoToItem.wdGoToPage;
        //        object Which1 = Word.WdGoToDirection.wdGoToNext;
        //        object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
        //        document.ActiveWindow.Selection.GoTo(ref What1, ref Which1, ref missing, ref Name); // 第二个参数可以用Nothing
        //        document.ActiveWindow.Selection.InsertBreak(ref oPageBreak);

        //        int pages = document.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages);//word总页数
        //        object limitePage = (object)(pages - WORD_BACK_NUM + 1);
        //        document.ActiveWindow.Selection.GoTo(ref What1, ref Which1, ref missing, ref limitePage); // 第二个参数可以用Nothing
        //        document.ActiveWindow.Selection.InsertBreak(ref oPageBreak);
        //        //document.Save();
        //        #endregion 在第WORD_FRONT_NUM页末尾插入分节符,在第WORD_BACK_NUM页前插入分节符,避免提取过多页数
        //        Utility.Log.TextLog.WritwLog("madeword4");
        //        #region 处理前n页,循环处理该word的section
        //        foreach (Word.Section section in document.Sections)
        //        {
        //            var range = section.Range;
        //            Utility.Log.TextLog.WritwLog("madeword4.1");
        //            if (range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) > WORD_FRONT_NUM)
        //                break;
        //            bool HasFetched = false;
        //            Utility.Log.TextLog.WritwLog("madeword4.2");
        //            for (int j = 1; j <= range.Paragraphs.Count; j++)
        //            {

        //                string text = section.Range.Paragraphs[j].Range.Text.Trim().Replace(" ", "").Replace("　", "");
        //                if (Front5WordFetch.Any(a => { return text.Contains(a); }))//如果该段文字包含配置的敏感字,则复制出来这个section
        //                {
        //                    Utility.Log.TextLog.WritwLog("madeword4.3");
        //                    section.Range.Copy();
        //                    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                    Utility.Log.TextLog.WritwLog("madeword4.4");
        //                    P_document.ActiveWindow.Selection.Paste();

        //                    HasFetched = true;
        //                    break;
        //                }
        //            }
        //            if (HasFetched)
        //                continue;
        //            //查找版式为嵌入式的图片,图片默认都是嵌入式
        //            foreach (Word.InlineShape inlShape in range.InlineShapes)
        //            {
        //                if (inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapePicture) || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
        //                    || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject))
        //                {
        //                    Utility.Log.TextLog.WritwLog("madeword4.5");
        //                    //inlShape.Range.Copy();
        //                    range.Copy();
        //                    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                    P_document.ActiveWindow.Selection.Paste();

        //                    Utility.Log.TextLog.WritwLog("madeword4.6");
        //                    break;
        //                }
        //            }
        //            //查找其他版式的图片
        //            //for (int i = 1; i <= range.ShapeRange.Count; i++)
        //            //{
        //            //    range.ShapeRange[i].Select();
        //            //    document.ActiveWindow.Selection.Copy();
        //            //    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //            //    P_document.ActiveWindow.Selection.Paste();
        //            //}
        //            if (range.ShapeRange.Count > 0)
        //            {
        //                range.Copy();
        //                P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                P_document.ActiveWindow.Selection.Paste();
        //            }

        //        }
        //        #endregion
        //        Utility.Log.TextLog.WritwLog("madeword5");
        //        #region 处理后5页

        //        if (pages > WORD_FRONT_NUM)//如果总页数大于已提取过的前n页,则从后几页开始提取
        //        {
        //            int secCount = document.Sections.Count;//section总数
        //            int xxCount = pages < WORD_FRONT_NUM + WORD_BACK_NUM ? pages - WORD_FRONT_NUM : WORD_BACK_NUM;//实际计算出的要寻找的后几页
        //            //如果最后一个section的起始页已经超出后n页的寻找范围,则找出其中的图片
        //            //if (document.Sections[secCount].Range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) <= pages - xxCount)
        //            //{

        //            //    bool copyed = false;
        //            //    Word.Range range = document.Sections[secCount].Range;


        //            //    string text = range.Text.Trim().Replace(" ", "").Replace("　", "");
        //            //    if (Back5Array.Any(a => { return text.Contains(a); }))
        //            //    {
        //            //        range.Copy();
        //            //        P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //            //        P_document.ActiveWindow.Selection.Paste();
        //            //        copyed = true;
        //            //    }

        //            //    if (!copyed)
        //            //    {
        //            //        //查找版式为嵌入式的图片,图片默认都是嵌入式
        //            //        foreach (Word.InlineShape inlShape in range.InlineShapes)
        //            //        {
        //            //            if (inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapePicture) || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
        //            //                || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject))
        //            //            {
        //            //                if (inlShape.Range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) > pages - xxCount)
        //            //                {
        //            //                    inlShape.Range.Copy();
        //            //                    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //            //                    P_document.ActiveWindow.Selection.Paste();
        //            //                }
        //            //            }
        //            //        }
        //            //        //查找其他版式的图片
        //            //        for (int i = 1; i <= range.ShapeRange.Count; i++)
        //            //        {
        //            //            range.ShapeRange[i].Select();
        //            //            document.ActiveWindow.Selection.Copy();
        //            //            P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //            //            P_document.ActiveWindow.Selection.Paste();
        //            //        }
        //            //    }
        //            //}
        //            //最后一个section的起始页没有超出后n页的寻找范围,则从后往前循环处理section
        //            //else
        //            //{
        //            for (int i = secCount; i > 0; i--)
        //            {
        //                var range = document.Sections[i].Range;
        //                var fs = range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber);
        //                if (range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) <= pages - xxCount)
        //                    break;
        //                string text = range.Text.Trim().Replace(" ", "").Replace("　", "");
        //                if (Back5Array.Any(a => { return text.Contains(a); }))
        //                {
        //                    range.Copy();
        //                    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                    P_document.ActiveWindow.Selection.Paste();

        //                }
        //                else
        //                {
        //                    //查找嵌入式版式的图片
        //                    foreach (Word.InlineShape inlShape in range.InlineShapes)
        //                    {
        //                        if (inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapePicture) || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
        //                            || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject))
        //                        {
        //                            inlShape.Range.Copy();
        //                            P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                            try
        //                            {
        //                                P_document.ActiveWindow.Selection.Paste();
        //                            }
        //                            catch
        //                            {
        //                                continue;
        //                            }
        //                        }
        //                    }
        //                    //查找其他版式的图片
        //                    for (int j = 1; j <= range.ShapeRange.Count; j++)
        //                    {
        //                        range.ShapeRange[j].Select();
        //                        document.ActiveWindow.Selection.Copy();
        //                        P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
        //                        try
        //                        {
        //                            P_document.ActiveWindow.Selection.Paste();
        //                        }
        //                        catch
        //                        {
        //                            continue;
        //                        }
        //                    }
        //                }
        //            }
        //            //}
        //        }
        //        #endregion
        //        Utility.Log.TextLog.WritwLog("madeword6");
        //        if (P_document.Words.Count > 1)//判断是否提取出了有效信息,如果没有则返回null
        //            P_document.ExportAsFixedFormat(fullPdfName, Word.WdExportFormat.wdExportFormatPDF);
        //        else
        //            fullPdfName = null;
        //        Utility.Log.TextLog.WritwLog("madeword7");
        //        P_document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        //        Utility.Log.TextLog.WritwLog("madeword8");
        //        document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        //        Utility.Log.TextLog.WritwLog("madeword9");
        //        _app.Quit();
        //        KillWord();
        //        Utility.Log.TextLog.WritwLog("madeword10");
        //        return fullPdfName;
        //    }
        //    catch (Exception e)
        //    {
        //        ShowLog("MadeWord异常:" + e.Message, 3);
        //        int linenum = new System.Diagnostics.StackFrame(true).GetFileLineNumber();
        //        Utility.Log.TextLog.WritwLog("整理" + filename + "失败:" + e.Message + ",错误行数:" + linenum, true);
        //        if (document != null)
        //            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        //        if (_app != null)
        //            _app.Quit();
        //        KillWord();
        //        return "";
        //    }
        //}
        //#endregion

        #region 提取word非正文页4.0
        string[] Front5WordFetch = ConfigHelper.GetValue("Front5DocFetchWord").Split(new char[] { ',' });
        public string MadeWord(object wordFilePath)
        {
            Utility.Log.TextLog.WritwLog("madeword1");
            string filename = System.IO.Path.GetFileName(wordFilePath.ToString());
            Utility.Log.TextLog.WritwLog("madeword2");
            Word.Application _app = null;
            Word.Document document = null;
            object missing = System.Reflection.Missing.Value;
            try
            {
                //整理后生成的文件名
                Utility.Log.TextLog.WritwLog("madeword3");
                string fullPdfName = ViewModel.Upload_Path + "\\" + filename.Substring(0, filename.LastIndexOf('.')) + "[整].pdf";
                Utility.Log.TextLog.WritwLog("madeword3-1" + wordFilePath);
                _app = new Word.Application();
                Utility.Log.TextLog.WritwLog("madeword3-2");
                document = _app.Documents.Open(wordFilePath.ToString(), false, false, false, ref missing, missing, false, missing, missing, missing, missing, false, false, missing, true, missing);
                Utility.Log.TextLog.WritwLog("madeword3-3");
                Word.Document P_document = _app.Documents.Add(ref missing, ref missing, ref missing);
                object What = Word.WdGoToItem.wdGoToSection;
                object Which = Word.WdGoToDirection.wdGoToLast;

                int pages = document.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages);//word总页数

                for (int i = 1; i <= pages; i++)
                {
                    Word.Range wrg1;
                    Word.Range wrg2;
                    Word.Range wrg;
                    wrg1 = document.GoTo(ref What, ref Which, i);
                    wrg2 = wrg1.GoToNext(Word.WdGoToItem.wdGoToPage);
                    wrg = document.Range(wrg1.Start, wrg2.Start);//指定页的范围
                    string strContent = wrg.Text;//获取该页内容
                    
                }
                
                #region 处理前n页,循环处理该word的section
                foreach (Word.Section section in document.Sections)
                {
                    var range = section.Range;
                    Utility.Log.TextLog.WritwLog("madeword4.1");
                    if (range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) > WORD_FRONT_NUM)
                        break;
                    bool HasFetched = false;
                    Utility.Log.TextLog.WritwLog("madeword4.2");
                    for (int j = 1; j <= range.Paragraphs.Count; j++)
                    {

                        string text = section.Range.Paragraphs[j].Range.Text.Trim().Replace(" ", "").Replace("　", "");
                        if (Front5WordFetch.Any(a => { return text.Contains(a); }))//如果该段文字包含配置的敏感字,则复制出来这个section
                        {
                            Utility.Log.TextLog.WritwLog("madeword4.3");
                            section.Range.Copy();
                            P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                            Utility.Log.TextLog.WritwLog("madeword4.4");
                            P_document.ActiveWindow.Selection.Paste();

                            HasFetched = true;
                            break;
                        }
                    }
                    if (HasFetched)
                        continue;
                    //查找版式为嵌入式的图片,图片默认都是嵌入式
                    foreach (Word.InlineShape inlShape in range.InlineShapes)
                    {
                        if (inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapePicture) || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
                            || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject))
                        {
                            Utility.Log.TextLog.WritwLog("madeword4.5");
                            //inlShape.Range.Copy();
                            range.Copy();
                            P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                            P_document.ActiveWindow.Selection.Paste();

                            Utility.Log.TextLog.WritwLog("madeword4.6");
                            break;
                        }
                    }
                 
                    if (range.ShapeRange.Count > 0)
                    {
                        range.Copy();
                        P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                        P_document.ActiveWindow.Selection.Paste();
                    }

                }
                #endregion
               
                #region 处理后5页

                if (pages > WORD_FRONT_NUM)//如果总页数大于已提取过的前n页,则从后几页开始提取
                {
                    int secCount = document.Sections.Count;//section总数
                    int xxCount = pages < WORD_FRONT_NUM + WORD_BACK_NUM ? pages - WORD_FRONT_NUM : WORD_BACK_NUM;//实际计算出的要寻找的后几页
                  
                    for (int i = secCount; i > 0; i--)
                    {
                        var range = document.Sections[i].Range;
                        var fs = range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber);
                        if (range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) <= pages - xxCount)
                            break;
                        string text = range.Text.Trim().Replace(" ", "").Replace("　", "");
                        if (Back5Array.Any(a => { return text.Contains(a); }))
                        {
                            range.Copy();
                            P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                            P_document.ActiveWindow.Selection.Paste();

                        }
                        else
                        {
                            //查找嵌入式版式的图片
                            foreach (Word.InlineShape inlShape in range.InlineShapes)
                            {
                                if (inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapePicture) || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
                                    || inlShape.Type.Equals(Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject))
                                {
                                    inlShape.Range.Copy();
                                    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                                    try
                                    {
                                        P_document.ActiveWindow.Selection.Paste();
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                            //查找其他版式的图片
                            for (int j = 1; j <= range.ShapeRange.Count; j++)
                            {
                                range.ShapeRange[j].Select();
                                document.ActiveWindow.Selection.Copy();
                                P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                                try
                                {
                                    P_document.ActiveWindow.Selection.Paste();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }
                #endregion
                Utility.Log.TextLog.WritwLog("madeword6");
                if (P_document.Words.Count > 1)//判断是否提取出了有效信息,如果没有则返回null
                    P_document.ExportAsFixedFormat(fullPdfName, Word.WdExportFormat.wdExportFormatPDF);
                else
                    fullPdfName = null;
                Utility.Log.TextLog.WritwLog("madeword7");
                P_document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Utility.Log.TextLog.WritwLog("madeword8");
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Utility.Log.TextLog.WritwLog("madeword9");
                _app.Quit();
                KillWord();
                Utility.Log.TextLog.WritwLog("madeword10");
                return fullPdfName;
            }
            catch (Exception e)
            {
                ShowLog("MadeWord异常:" + e.Message, 3);
                int linenum = new System.Diagnostics.StackFrame(true).GetFileLineNumber();
                Utility.Log.TextLog.WritwLog("整理" + filename + "失败:" + e.Message + ",错误行数:" + linenum, true);
                if (document != null)
                    document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                if (_app != null)
                    _app.Quit();
                KillWord();
                return "";
            }
        }
        #endregion

        #region 提取pdf非正文页1.0
        //public string MadePdf( string pdfFilePath)
        //{
        //    try
        //    {
        //        /*加密的pdf文件程序无法识别,但可以人工打开浏览*/
        //        PdfReader testReader = new PdfReader(pdfFilePath);
        //        if (testReader.IsEncrypted()) 
        //        {
        //            testReader.Close();
        //            return "";
        //        }
        //        testReader.Close();
        //        /*加密的pdf文件程序无法识别,但可以人工打开浏览*/
        //        string pdfname = System.IO.Path.GetFileName(pdfFilePath);
        //        string fullPdfName = ViewModel.Upload_Path + "\\" + pdfname.Substring(0, pdfname.LastIndexOf('.')) + "[整].pdf";
        //        PDF.Document document = new PDF.Document();
        //        PdfCopy writer = new PdfCopy(document, new FileStream(fullPdfName, FileMode.Create));
        //        document.Open();

        //        PdfReader reader = new PdfReader(pdfFilePath);
        //        reader.ConsolidateNamedDestinations();
        //        int pageNum = Convert.ToInt32(ConfigHelper.GetValue("PDF_ExNum"));
        //        pageNum = reader.NumberOfPages > pageNum ? pageNum : reader.NumberOfPages;
        //        for (int i = 1; i <= pageNum; i++)
        //        {
        //            PdfImportedPage page = writer.GetImportedPage(reader, i);
        //            writer.AddPage(page);
        //        }
        //        reader.Close();
        //        writer.Close();
        //        document.Close();
        //        return fullPdfName;
        //    }
        //    catch (Exception e)
        //    {
        //        Utility.Log.TextLog.WritwLog("整理" + pdfFilePath + "失败:" + e.Message, true);
        //        return "";
        //    }
        //}
        #endregion

        #region 提取pdf非正文页2.0
        /// <summary>
        /// 提取pdf非正文页2.0
        /// </summary>
        /// <param name="pdfFilePath">待提取pdf路径</param>
        /// <returns>"":异常 null:没有提取出有效信息</returns>
        public string MadePdf(string pdfFilePath)
        {
            #region 增加文件异常的验证,如果读取不了字则意味着文件打开异常 edit on 2016.06.16
            Reader reader = new Reader(1, 5);
            reader.ReadHandler = reader.ReadPdf;
            string header = reader.ReadWithTimeout(pdfFilePath);
            if (header == "文件读取异常" || header == "文件读取超时" || header == "文件内容为乱码" || header == "")
                return "";
            #endregion

            string pdfname = System.IO.Path.GetFileName(pdfFilePath);
            Document document = new Document();
            string fullPdfName = ViewModel.Upload_Path + "\\" + pdfname.Substring(0, pdfname.LastIndexOf('.')) + "[整].pdf";
            bool HAVE_CONTENT = false;//是否提取出有效信息
            try
            {
                PdfReader pdfReader = new PdfReader(pdfFilePath);
                //var ss=pdfReader.IsEncrypted();
                PdfCopy writer = new PdfCopy(document, new FileStream(fullPdfName, FileMode.Create));
                document.Open();
                int pagecount = pdfReader.NumberOfPages;
                int exCount = pagecount < PDF_FRONT_NUM ? pagecount : PDF_FRONT_NUM;

                //循环前5页(或小于5页),含有摘要的页就不提取
                for (int page = 1; page <= exCount; page++)
                {
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page).Replace(" ", "").Replace("　", "");
                    //如果整页只有图片没有文字,则提取
                    //if (PdfHavePic(pdfReader, page) && string.IsNullOrWhiteSpace(currentText))
                    if (PdfHavePic(pdfReader, page))
                    {
                        writer.AddPage(writer.GetImportedPage(pdfReader, page));
                        HAVE_CONTENT = true;
                        continue;
                    }


                    //currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    if (string.IsNullOrEmpty(currentText) || Front5Array.Any(a => { return currentText.Contains(a); }))
                        continue;
                    writer.AddPage(writer.GetImportedPage(pdfReader, page));
                    HAVE_CONTENT = true;
                }
                if (pagecount > PDF_FRONT_NUM)
                {
                    int xxCount = pagecount < PDF_FRONT_NUM + PDF_BACK_NUM ? pagecount - PDF_FRONT_NUM : PDF_BACK_NUM;
                    //循环后5页,有图片的页就提取(如果pdf总页数大于5小于10,则循环除前5页外的页数)
                    for (int page = pagecount - xxCount + 1; page <= pagecount; page++)
                    {
                        //先看这页有没有包含版权字段
                        string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page).Replace(" ", "").Replace("　", "");
                        //currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                        if (Back5Array.Any(a => { return currentText.Contains(a); }))
                        {
                            writer.AddPage(writer.GetImportedPage(pdfReader, page));
                            HAVE_CONTENT = true;
                            continue;
                        }
                        //再看这页是否只含图片,不含文字
                        //if (string.IsNullOrWhiteSpace(currentText)&&PdfHavePic(pdfReader, page))
                        if (PdfHavePic(pdfReader, page))
                        {
                            writer.AddPage(writer.GetImportedPage(pdfReader, page));
                            HAVE_CONTENT = true;
                            continue;
                        }
                    }
                }
                if (!HAVE_CONTENT)
                    fullPdfName = null;
                else
                {
                    writer.Close();
                    document.Close();
                }
                pdfReader.Close();

                return fullPdfName;
            }
            catch (Exception e)
            {
                Utility.Log.TextLog.WritwLog("MadePdf异常:" + e.Message);
                return "";
            }
        }
        #endregion

        /// <summary>
        /// 测试PDF文件的某页是否含有图片
        /// </summary>
        /// <param name="pdfReader">PdfReader实例</param>
        /// <param name="page">页码</param>
        /// <returns></returns>
        private bool PdfHavePic(PdfReader pdfReader, int page)
        {
            PdfDictionary pg = pdfReader.GetPageN(page);
            PdfDictionary res = (PdfDictionary)PdfReader.GetPdfObject(pg.Get(PdfName.RESOURCES));
            if (res != null)
            {
                PdfDictionary xobj = (PdfDictionary)PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT));
                if (xobj != null)
                {
                    foreach (PdfName name in xobj.Keys)
                    {
                        PdfObject obj = xobj.Get(name);
                        if (obj.IsIndirect())
                        {
                            PdfDictionary tg = (PdfDictionary)PdfReader.GetPdfObject(obj);
                            PdfName type = (PdfName)PdfReader.GetPdfObject(tg.Get(PdfName.SUBTYPE));
                            if (PdfName.IMAGE.Equals(type) || PdfName.FORM.Equals(type))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }
        #region pdf加水印
        /// <summary>
        /// 添加普通偏转角度文字水印
        /// </summary>
        /// <param name="inputfilepath">输入pdf文件</param>
        /// <param name="outputfilepath">输出pdf文件</param>
        /// <param name="waterMarkName">水印文字</param>
        /// <param name="page">第几页</param>
        public void SetWatermark(string inputfilepath, string outputfilepath, string waterMarkName, int page)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(inputfilepath);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                //int total = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\SIMFANG.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState();

                //content = pdfStamper.GetOverContent(i);//在内容上方加水印
                content = pdfStamper.GetUnderContent(page);//在内容下方加水印
                //透明度
                gs.FillOpacity = 0.8f;
                content.SetGState(gs);
                //content.SetGrayFill(0.3f);
                //开始写入文本
                content.BeginText();
                content.SetColorFill(BaseColor.RED);
                content.SetFontAndSize(font, 50);
                content.SetTextMatrix(0, 0);
                content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2, height - 50, 0);
                content.EndText();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
        }

        public void SetWatermark(string inputfilepath, string outputfilepath, System.Data.DataTable table)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(inputfilepath);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                //int total = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\SIMFANG.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState();
                foreach (System.Data.DataRow row in table.Rows)
                {
                    //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                    if (row[1].ToString() == "0" || row[2].ToString() == "0")
                        continue;
                    content = pdfStamper.GetOverContent(Convert.ToInt32(row[1]));//在内容上方加水印
                    //透明度
                    gs.FillOpacity = 0.8f;
                    content.SetGState(gs);
                    //content.SetGrayFill(0.3f);
                    //开始写入文本
                    content.BeginText();
                    content.SetColorFill(BaseColor.RED);
                    content.SetFontAndSize(font, 50);
                    content.SetTextMatrix(0, 0);
                    content.ShowTextAligned(Element.ALIGN_CENTER, row[0] + "任务开始", width / 2, height - 50, 0);
                    content.EndText();

                    content = pdfStamper.GetOverContent(Convert.ToInt32(row[2]));//在内容上方加水印
                    //透明度
                    gs.FillOpacity = 0.8f;
                    content.SetGState(gs);
                    //content.SetGrayFill(0.3f);
                    //开始写入文本
                    content.BeginText();
                    content.SetColorFill(BaseColor.RED);
                    content.SetFontAndSize(font, 50);
                    content.SetTextMatrix(0, 0);
                    content.ShowTextAligned(Element.ALIGN_CENTER, row[0] + "任务结束", width / 2, 20, 0);
                    content.EndText();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
        }
        #endregion

        #region 合并pdf
        /// <summary>
        /// 合并两个PDF文件
        /// </summary>
        /// <param name="firstFileName">第一个pdf文件名</param>
        /// <param name="sencondFileName">第二个pdf文件名</param>
        /// <param name="outFile">合并后文件名</param>
        /// <param name="start">第二个文件在合并后文件的起始页数</param>
        /// <param name="end">第二个文件在合并后文件的结束页数</param>
        public void CombineMultiplePDFs(string firstFileName, string sencondFileName, string outFile, out int start, out int end)
        {
            start = end = 0;
            string[] fileNames = { firstFileName, sencondFileName };
            PDF.Document document = new PDF.Document();
            PdfCopy writer = new PdfCopy(document, new FileStream(outFile, FileMode.Create));
            document.Open();
            for (int j = 0; j < fileNames.Length; j++)
            {
                PdfReader reader = new PdfReader(fileNames[j]);
                if (j == 0)
                {
                    start = reader.NumberOfPages + 1;
                }
                else if (j == 1)
                {
                    end = start + reader.NumberOfPages - 1;
                }
                reader.ConsolidateNamedDestinations();
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    PdfImportedPage page = writer.GetImportedPage(reader, i);
                    writer.AddPage(page);
                }
                reader.Close();
            }
            writer.Close();
            document.Close();
        }
        /// <summary>
        /// 合并多个pdf文件
        /// </summary>
        /// <param name="Paths">pdf路径集合</param>
        /// <param name="outFile">合并后的文件路径</param>
        /// <returns></returns>
        public bool CombineMultiplePDFs(List<string> Paths, ref string outFile)
        {
            try
            {
                PDF.Document document = new PDF.Document();
                PdfCopy writer = new PdfCopy(document, new FileStream(outFile, FileMode.Create));
                document.Open();
                for (int j = 0; j < Paths.Count; j++)
                {
                    if (Paths[j] == "不需" || Paths[j] == null)
                        continue;
                    PdfReader reader = new PdfReader(Paths[j]);
                    reader.ConsolidateNamedDestinations();
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        PdfImportedPage page = writer.GetImportedPage(reader, i);
                        writer.AddPage(page);
                    }
                    reader.Close();
                }
                writer.Close();
                document.Close();
                foreach (string s in Paths)
                {
                    if (s != "" && s != null)
                        File.Delete(s);
                }
                return true;
            }
            catch (Exception e)
            {
                Utility.Log.TextLog.WritwLog("合并pdf失败:" + e.Message, true);
                outFile = "";
                return false;

            }
        }
        #endregion

        /// <summary>
        /// 复制文件到目标文件夹，如果存在同名文件，则文件名前面加数字
        /// </summary>
        /// <param name="file">文件</param>
        /// <param name="des">目标文件夹路径</param>
        /// 返回实际复制后的文件名
        public string CopyFile(FileInfo file, string des)
        {
            string fileName = file.Name;
            int i = 1;
            while (File.Exists(Path.Combine(des, fileName)))
            {
                fileName = file.Name.Remove(file.Name.LastIndexOf('.')) + "_" + i.ToString() + file.Extension;
                i++;
            }
            File.Copy(file.FullName, Path.Combine(des, fileName));
            return fileName;
        }

        /// <summary>
        /// 在本地目录生成固定提示pdf文件
        /// </summary>
        public bool Voidpdf()
        {
            try
            {
                //初始化一个目标文档类 
                Document document = new Document(PageSize.A4);
                //调用PDF的写入方法流
                //注意FileMode-Create表示如果目标文件不存在，则创建，如果已存在，则覆盖。
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(TEMP_PDF, FileMode.Create));
                document.Open();
                PdfContentByte cb = writer.DirectContent;
                cb.BeginText();
                BaseFont bfont = BaseFont.CreateFont(@"c:\windows\fonts\SIMHEI.TTF", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);//设定字体：黑体 
                cb.SetFontAndSize(bfont, 34);//设定字号 
                cb.SetCharacterSpacing(1);//设定字间距 
                cb.SetRGBColorFill(66, 00, 00);//设定文本颜色 
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "没有提取出有效信息", PageSize.A4.Width / 2, PageSize.A4.Height / 2 + 20, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "请打开源文件查看", PageSize.A4.Width / 2, PageSize.A4.Height / 2 - 20, 0);
                cb.EndText();
                document.Close();
                //关闭写入流
                writer.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void KillWord()
        {
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
            {
                p.Kill();
            }
        }

        /// <summary>
        /// 按用户配置直接提取固定页数
        /// </summary>
        /// <param name="pdfFilePath">pdf路径</param>
        /// <param name="frontPage">前几页</param>
        /// <param name="backPage">后几页</param>
        /// <returns></returns>
        public string MadeDefinedPdf(string pdfFilePath, int frontPage, int backPage)
        {
            string pdfname = System.IO.Path.GetFileName(pdfFilePath);
            Document document = new Document();
            string fullPdfName = ViewModel.Upload_Path + "\\" + pdfname.Substring(0, pdfname.LastIndexOf('.')) + "[整].pdf";
            try
            {
                PdfReader pdfReader = new PdfReader(pdfFilePath);
                PdfCopy writer = new PdfCopy(document, new FileStream(fullPdfName, FileMode.Create));
                document.Open();
                int pagecount = pdfReader.NumberOfPages;
                for (int page = 1; page <= frontPage; page++)
                {
                    writer.AddPage(writer.GetImportedPage(pdfReader, page));
                }
                for (int page = pagecount - backPage + 1; page <= pagecount; page++)
                {
                    writer.AddPage(writer.GetImportedPage(pdfReader, page));
                }
                writer.Close();
                document.Close();
                pdfReader.Close();

                return fullPdfName;
            }
            catch (Exception e)
            {
                Utility.Log.TextLog.WritwLog("定制MadePdf异常:" + e.Message);
                return "";
            }
        }
        /// <summary>
        /// 按用户配置直接提取固定页数
        /// </summary>
        /// <param name="pdfFilePath">Word路径</param>
        /// <param name="frontPage">前几页</param>
        /// <param name="backPage">后几页</param>
        public string MadeDefinedWord111111111(object prams)
        {
            DefinedWordPra wordpra = prams as DefinedWordPra;
            string wordFilePath=wordpra.path;
            int frontPage=wordpra.front;
            int backPage = wordpra.back;
            string filename = System.IO.Path.GetFileName(wordFilePath.ToString());
            Word.Application _app = null;
            Word.Document document = null;
            object missing = System.Reflection.Missing.Value;
            string fullPdfName = ViewModel.Upload_Path + "\\" + filename.Substring(0, filename.LastIndexOf('.')) + "[整].pdf";
            _app = new Word.Application();
            document = _app.Documents.Open(wordFilePath.ToString(), false, false, false, ref missing, missing, false, missing, missing, missing, missing, false, false, missing, true, missing);
            Word.Document P_document = _app.Documents.Add(ref missing, ref missing, ref missing);
            object What = Word.WdGoToItem.wdGoToSection;
            object Which = Word.WdGoToDirection.wdGoToLast;
            try
            {
                #region 在第WORD_FRONT_NUM页末尾插入分节符,在第WORD_BACK_NUM页前插入分节符,避免提取过多页数
                object Name = (object)(frontPage + 1);
                object What1 = Word.WdGoToItem.wdGoToPage;
                object Which1 = Word.WdGoToDirection.wdGoToNext;
                object oPageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
                document.ActiveWindow.Selection.GoTo(ref What1, ref Which1, ref missing, ref Name); // 第二个参数可以用Nothing
                document.ActiveWindow.Selection.InsertBreak(ref oPageBreak);
                int pages = document.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages);//word总页数
                if (backPage > 0)
                {
                    
                    object limitePage = (object)(pages - backPage + 1);
                    document.ActiveWindow.Selection.GoTo(ref What1, ref Which1, ref missing, ref limitePage); // 第二个参数可以用Nothing
                    document.ActiveWindow.Selection.InsertBreak(ref oPageBreak);
                }
                document.Save();
                #endregion 在第WORD_FRONT_NUM页末尾插入分节符,在第WORD_BACK_NUM页前插入分节符,避免提取过多页数

                #region 提取前n页
                if (frontPage > 0)
                {
                    foreach (Word.Section section in document.Sections)
                    {
                        var range = section.Range;
                        var text = range.Text;
                        var cccc = range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber);
                        var ffff = range.Characters[range.Characters.Count].get_Information(Word.WdInformation.wdActiveEndPageNumber);
                        if (range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) > frontPage)
                            break;
                        section.Range.Copy();
                        P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                        P_document.ActiveWindow.Selection.Paste();
                    }
                }
                #endregion

                #region 提取后n页
                if (backPage > 0)
                {
                    int secCount = document.Sections.Count;//section总数
                    for (int i = secCount; i > 0; i--)
                    {
                        var range = document.Sections[i].Range;
                        if (range.Characters[1].get_Information(Word.WdInformation.wdActiveEndPageNumber) <= pages - backPage)
                            break;
                        range.Copy();
                        P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                        P_document.ActiveWindow.Selection.Paste();
                    }
                }
                #endregion
                P_document.ExportAsFixedFormat(fullPdfName, Word.WdExportFormat.wdExportFormatPDF);
                P_document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                _app.Quit();
                return fullPdfName;
            }
            catch (Exception e)
            {
                ShowLog("定制MadeWord异常:" + e.Message, 3);
                if (document != null)
                    document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                if (_app != null)
                    _app.Quit();
                KillWord();
                return "";
            }


        }

        public string MadeDefinedWord(object prams)
        {
            DefinedWordPra wordpra = prams as DefinedWordPra;
            string wordFilePath = wordpra.path;
            int frontPage = wordpra.front;
            int backPage = wordpra.back;

            object missing = System.Reflection.Missing.Value;
            object What = Word.WdGoToItem.wdGoToSection;
            object Which = Word.WdGoToDirection.wdGoToLast;

            Word.Application _app = null;
            Word.Document document = null;
            try
            {
                string filename = System.IO.Path.GetFileName(wordFilePath.ToString());
                string fullPdfName = ViewModel.Upload_Path + "\\" + filename.Substring(0, filename.LastIndexOf('.')) + "[整].pdf";
                string fullFrontName = ViewModel.Upload_Path + "\\" + filename.Substring(0, filename.LastIndexOf('.')) + "[整前].pdf";
                string fullBackName = ViewModel.Upload_Path + "\\" + filename.Substring(0, filename.LastIndexOf('.')) + "[整后].pdf";

                _app = new Word.Application();
                document = _app.Documents.Open(wordFilePath.ToString(), false, false, false, ref missing, missing, false, missing, missing, missing, missing, false, false, missing, true, missing);

                int pages = document.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages);//word总页数
                if (frontPage > pages)
                    frontPage = pages;

                document.ExportAsFixedFormat(fullFrontName, Word.WdExportFormat.wdExportFormatPDF, false, Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                    Word.WdExportRange.wdExportFromTo, 1, frontPage, Word.WdExportItem.wdExportDocumentWithMarkup, false, true, Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false, Type.Missing);

                if (backPage > 0)
                {
                    if (backPage > pages)
                        backPage = pages;
                    document.ExportAsFixedFormat(fullBackName, Word.WdExportFormat.wdExportFormatPDF, false, Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                    Word.WdExportRange.wdExportFromTo, pages - backPage + 1, pages, Word.WdExportItem.wdExportDocumentWithMarkup, false, true, Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false, Type.Missing);
                    List<string> list=new List<string>();
                    list.Add(fullFrontName);
                    list.Add(fullBackName);
                    CombineMultiplePDFs(list, ref fullPdfName);
                }
                else
                {
                    System.IO.File.Move(fullFrontName, fullPdfName);
                }


                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                _app.Quit();
                return fullPdfName;
            }
            catch (Exception e)
            {
                ShowLog("定制MadeWord异常:" + e.Message, 3);
                if (document != null)
                    document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                if (_app != null)
                    _app.Quit();
                KillWord();
                return "";
            }

        }
    }

    public class DefinedWordPra
    {
        public string path { get; set; }
        public int front { get; set; }
        public int back { get; set; }
    }

}
