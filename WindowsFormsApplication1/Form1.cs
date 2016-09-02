using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string DIR;

        public Form1()
        {
            InitializeComponent();
            textBox1.Text=@"D:\work\51882\20160122\upload";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DirectoryInfo dir = new DirectoryInfo(textBox1.Text);
            DIR = dir.FullName;
            foreach (FileInfo file in dir.GetFiles())
            { 
                listView1.Items.Add(new ListViewItem(file.Name));
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                int selectCount = listView1.SelectedItems.Count;
                if (selectCount > 0)
                {
                    string filename = listView1.SelectedItems[0].Text;
                    var sfsf = Path.Combine(DIR, filename);
                    axAcroPDF1.src = sfsf;
                    //bool fsf = axAcroPDF1.LoadFile(sfsf);

                    using (FileStream pdfFileStream = new FileStream(sfsf, FileMode.Open, FileAccess.Read))
                    {
                        // load PDF from stream
                        pdfViewer1.LoadStream(pdfFileStream);
                        //lblPageCount.Text = pdfViewerControl.PageCount > 0 ? String.Format("Page Count: {0}", pdfViewerControl.PageCount) : String.Empty;
                    }
                   
                }
            }
            catch (Exception ee)
            { }
        }
    }
}
