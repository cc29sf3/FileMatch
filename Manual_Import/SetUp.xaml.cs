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

namespace Manual_Import
{
    /// <summary>
    /// SetUp.xaml 的交互逻辑
    /// </summary>
    public partial class SetUp : Window
    {
        public Action<int, int, int, int> ConfirmDifine;
        public Action CancalDifine;

        public SetUp()
        {
            InitializeComponent(); 
        }
        public SetUp(bool isDefined):this()
        {
            checkBox.IsChecked = isDefined;
        }
        public SetUp(int pdfFront, int pdfBack, int wordFront, int wordBack) : this(true)
        {
            this.pdfFront.Text = pdfFront.ToString();
            this.pdfBack.Text = pdfBack.ToString();
            this.wordFront.Text = wordFront.ToString();
            this.wordBack.Text = wordBack.ToString();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (checkBox.IsChecked.Value)
            {
                int a,b,c,d;
                if (int.TryParse(pdfFront.Text, out a) && int.TryParse(pdfBack.Text, out b) && int.TryParse(wordFront.Text, out c) && int.TryParse(wordBack.Text, out d))
                    ConfirmDifine(a, b, c, d);
                else
                {
                    this.Title += ":参数设置有误";
                    return;
                }
            }
            else
            {
                CancalDifine();
            }
            DialogResult = true;
            this.Close();
        }
    }
}
