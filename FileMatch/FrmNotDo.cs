using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FileMatch
{
    public partial class FrmNotDo : Form
    {
        public string info { get; set; }
        public FrmNotDo()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            info = string.IsNullOrWhiteSpace(textBox1.Text) ? comboBox1.Text : textBox1.Text;
            DialogResult=DialogResult.OK;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}
