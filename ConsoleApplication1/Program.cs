#define DEBUG
#define MYTEST

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
//using NPOI.XWPF;
//using NPOI.XWPF.UserModel;
using System.IO;
using System.Xml;
using SharpCompress;
using SharpCompress.Reader;
using SharpCompress.Common;
using SharpCompress.Archive;

using iTextSharp.text;

using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using iTextSharp.text.pdf;
using System.Xml.Linq;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Security.Cryptography;
using System.Collections;
using System.Text.RegularExpressions;
//using NPOI.XWPF;
//using POI=NPOI.XWPF.UserModel;
//using NPOI;
//using NPOI.POIFS.FileSystem;

namespace ConsoleApplication1
{
    #region//重写父类同名方法
    class MyClass
    {
        public void ShowMessage()
        { Console.WriteLine("我是父类"); }
    }

    class MyClass1 : MyClass
    {
        public void ShowMessage() // 这里就发生了重写,也可以说是隐藏了父类的方法. 这样做了之后就不能再使用父类的同名方法了.
        {
            Console.WriteLine("我是子类");
            base.ShowMessage();
        }
    }
    #endregion

    #region 覆盖父类的同名方法.
    class My
    {
        public virtual void SHowMessage()  //使用virtual关键字表示这个方法会被覆盖.
        { Console.WriteLine("我是父类,我将要被覆盖."); }
    }

    class My1 : My
    {
        public override void SHowMessage()  // 使用override 关键字来表示覆盖父类的同名方法.  覆盖和重写不同的是覆盖可以再调用父类的同名方法, 加一个base关键字就可以了.
        {
            Console.WriteLine("我是子类,我覆盖了父类的同名方法");
            base.SHowMessage();  // 这里就调用了父类的SHowMessage方法.
        }
    }
    #endregion


    class Transport
    {
        public string Name { get; set; }
        public string How { get; set; }
        public Transport(string n, string h)
        {
            Name = n;
            How = h;
        }
    }
    class vehicle
    {
        public string vehicleName { get; set; }
        public string vehicleHow { get; set; }
        public vehicle(string n, string h)
        {
            vehicleName = n;
            vehicleHow = h;
        }
    }

    class Program
    {




        static void Main(string[] args)
        {
            Word.Application _app = null;
            Word.Document document = null;
            object missing = System.Reflection.Missing.Value;
            _app = new Word.Application();
            string wordFilePath = @"D:\work\53002\20160823\131163876229304835\6\2062\20160822003\中共陕西省委党校2016年论文\2013210084-陈伟丽.doc";
            document = _app.Documents.Open(wordFilePath, false, false, false, ref missing, missing, false, missing, missing, missing, missing, false, false, missing, true, missing);
            object What = Word.WdGoToItem.wdGoToSection;
            object Which = Word.WdGoToDirection.wdGoToNext;

            Word.WdStatistic staticword = Word.WdStatistic.wdStatisticPages;
            int ipagecount = document.ComputeStatistics(staticword);//获得word文档的页数

            //跳转到指定的页数
            object pWhat = Word.WdGoToItem.wdGoToPage;
            object pWhich = Word.WdGoToDirection.wdGoToAbsolute;

            Word.Document P_document = _app.Documents.Add(ref missing, ref missing, ref missing);
            for (int i = 1; i <= ipagecount; i++)
            {
                if (i > 10)
                    break;
                Word.Range wrg1;
                Word.Range wrg2;
                Word.Range wrg;
                wrg1 = document.GoTo(ref pWhat, ref pWhich, i);
                wrg2 = wrg1.GoToNext(Word.WdGoToItem.wdGoToPage);
                wrg = document.Range(wrg1.Start, wrg2.Start);//指定页的范围

                string strContent = wrg.Text;//获取该页内容

                Console.WriteLine(strContent);

                if (strContent.Contains("答辩日期"))
                {
                  
                    wrg.Copy();
                    P_document.ActiveWindow.Selection.GoTo(ref What, ref Which, ref missing, ref missing);
                    P_document.ActiveWindow.Selection.Paste();
                }
            }
            P_document.ExportAsFixedFormat(@"D:\work\53002\20160823\131163876229304835\6\2062\20160822003\中共陕西省委党校2016年论文\xxxx.pdf", Word.WdExportFormat.wdExportFormatPDF);
            Console.WriteLine("over");
            P_document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            _app.Quit();
            Console.ReadKey();
        }






    }
}
