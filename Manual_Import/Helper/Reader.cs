using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
//using iTextSharp.text;
//using iTextSharp.text.pdf;
//using iTextSharp.text.pdf.parser;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Text.RegularExpressions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace Manual_Import.Helper
{
    /// <summary>
    /// 读取文件摘要类，包括读word方法和读pdf方法
    /// </summary>
    class Reader
    {
        public Action<string> ReadHandler;//读文件委托
        private ManualResetEvent mTimeoutObject;//多线程阻塞控制对象
        string text;//最后读取出的字符串
        //string jieshowNo;//接收编号
        int readWordCount;//要读取word前几个字符
        int readPdfPages;//读取pdf前几页

       
        object oMissing = System.Reflection.Missing.Value;
        object What = Word.WdGoToItem.wdGoToSection;
        object Which = Word.WdGoToDirection.wdGoToFirst;
        //string P_str_path = @"D:\\非正文页.doc";
       
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="jieshow">接收编号</param>
        public Reader(int word,int pdf)
        {
            
            readWordCount = word;
            readPdfPages = pdf;
            mTimeoutObject = new ManualResetEvent(true);
           
        }

        public void ReadWord(string path)
        {
            text = "";// path.Substring(path.LastIndexOf('\\') + 1);
            Word.Application app = new Word.Application();
            Word.Document doc = null;

            Word.Range range = null;
            try
            {
                Utility.Log.TextLog.WritwLog(path + "开始读取正常字符", false);
                object name = path as object;
                doc = app.Documents.Open(name, false, false, false, ref oMissing, oMissing, false, oMissing, oMissing, oMissing, oMissing, false, false, oMissing, true, oMissing);
                object start = 0, end = readWordCount;
                try//读取前xx个字符
                {
                    range = doc.Range(ref start, ref end);
                    text = text + range.Text;
                }
                catch (Exception e)//如果失败很可能由于没有那么多字符，就读取所有字符
                { text = text + doc.Content.Text; }

            }
            catch (Exception e)
            {
                text = "文件读取异常";
                Utility.Log.TextLog.WritwLog(path + "读取正常字符出错", true);
                Utility.Log.TextLog.WritwLog(e.Message);
                if (doc != null)
                {
                    doc.Close();
                    app.Quit();
                }

                return;
            }
            //读取文本框里的文本..纵向文本框会发生异常阻塞..暂时先注释掉
            try
            {
                //Utility.Log.TextLog.WritwLog("开始读取文本框", false);
                //var shaps = range.ShapeRange;
                //foreach (Word.Shape shap in shaps)
                //{
                //    if (shap.TextFrame.HasText != 0)
                //    {
                //        text += shap.TextFrame.TextRange.Text;
                //    }
                //}
            }
            catch (Exception ww)
            {
                Utility.Log.TextLog.WritwLog("读取文本框出错", true);
            }
            text = text.Replace("", "").Replace(" ", "").Replace(" ", "").Replace("　", "").Replace("\r", "").Replace("\f", "").Replace("\n", "").Replace("\t", "").Replace(" ", "").Replace("\a", "");
            if (!Regex.IsMatch(text, "[\u3E00-\u9FA5]"))
            {
                text = "文件内容为乱码";
                Utility.Log.TextLog.WritwLog(path+":文件内容为乱码", true);
            }

            try
            {
                Utility.Log.TextLog.WritwLog("开始关闭word", false);
                doc.Close(ref oMissing, ref oMissing, ref oMissing);
                app.Quit(ref oMissing, ref oMissing, ref oMissing);
            }
            catch (Exception hh)
            {
                Utility.Log.TextLog.WritwLog("关闭word出错", true);
                return;
            }


        }

        /// <summary>
        /// 用北京给的pdf工具读文件
        /// </summary>
        /// <param name="path">文件路径</param>
        public void ReadPdf(string path)
        {
            text = "";
            FileStream fs = null;
            StreamReader sr = null;
            Process proc = null;
            try
            {

                FileStream fs1 = new FileStream("1.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                fs1.SetLength(0);
                fs1.Close();


                //-f 1 -l 3意思是读取前三页，编码必须制定GBK，否则某些文件读不出汉字
                string arg = "-enc GBK -f 1 -l " + readPdfPages + " \"" + path + "\" \"1.txt\"";
                ProcessStartInfo info = new ProcessStartInfo("pdftotext.exe", arg);
                info.UseShellExecute = false;
                info.RedirectStandardInput = true;
                info.RedirectStandardOutput = true;
                info.RedirectStandardError = true;
                info.CreateNoWindow = true;
                proc = Process.Start(info);
                proc.WaitForExit();
                fs = new FileStream("1.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                sr = new StreamReader(fs, Encoding.GetEncoding("GBK"));

                text = sr.ReadToEnd();
                text = text.Replace("", "").Replace(" ", "").Replace(" ", "").Replace("　", "").Replace("\r", "").Replace("\f", "").Replace("\n", "").Replace("\t", "").Replace(" ", "").Replace("\a", "");
                if (string.IsNullOrEmpty(text))
                    text = ReadPdf2(path);
                text = text.Replace("", "").Replace(" ", "").Replace(" ", "").Replace("　", "").Replace("\r", "").Replace("\f", "").Replace("\n", "").Replace("\t", "").Replace(" ", "").Replace("\a", "");
            }
            catch (Exception ef)
            {

                text = ReadPdf2(path);
            }
            finally
            {
                if (sr != null)
                    sr.Close();
                if (fs != null)
                    fs.Close();
            }
        }

        ///原有读pdf的方法，已作废
        public string ReadPdf2(string path)
        {
            text = "";
            PdfReader pr = null;
            try
            {

                pr = new PdfReader(path);
                PdfReaderContentParser parser = new PdfReaderContentParser(pr);
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                int i, Page = 4;
                if (Page > pr.NumberOfPages)
                    Page = pr.NumberOfPages;
                for (i = 1; i <= Page; i++)
                    text += PdfTextExtractor.GetTextFromPage(pr, i, strategy);
                pr.Close();

                text = text.Replace("", "").Replace(" ", "").Replace(" ", "").Replace("　", "").Replace("\r", "").Replace("\f", "").Replace("\n", "").Replace("\t", "").Replace(" ", "");
                if (text.Length > 300)
                    text = text.Substring(0, 300);
            }
            catch (Exception e)
            {
                if (pr != null)
                {
                    pr.Close();
                }
                text = "文件读取异常";
            }
            return text;
        }

        /// <summary>
        /// 读取文件摘要，主线程阻塞60秒
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns>文件摘要字符串</returns>
        public string ReadWithTimeout(string path)
        {
            if (this.ReadHandler == null)
            {
                return "";
            }
            this.text = "";
            mTimeoutObject.Reset();
            ReadHandler.BeginInvoke(path, DoAsyncCallBack, null);
            //mTimeoutObject.WaitOne();
            if (!this.mTimeoutObject.WaitOne(new TimeSpan(0, 0, 60), false))
            {
                KillWord();
                return "文件读取超时";
            }
            return this.text;
        }

        /// <summary>
        /// ReadHandler回调函数
        /// </summary>
        private void DoAsyncCallBack(IAsyncResult result)
        {
            try
            {
               // ReadHandler.EndInvoke(result);极端情况下这部会出错，如超时时间设置为1秒，让文件连续超时
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                mTimeoutObject.Set();
            }
        }

        /// <summary>
        /// 杀掉所有word进程
        /// </summary>
        public static void KillWord()
        {
            foreach (Process p in Process.GetProcessesByName("pdftotext"))
            {
                p.Kill();
            }
            
            foreach (Process p in Process.GetProcessesByName("WINWORD"))
            {
                p.Kill();
            }
        }

       
    }
}
