using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Setting = WpfApp1.Properties.Settings;
namespace WpfApp1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            if (true)
            {
                //
            }
            BtnOpen_Click(new object(),new RoutedEventArgs());
            foreach (var item in Setting.Default.a)
            {
                LstMain.Items.Add(item);
            }
           
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.Filter="演示文稿|*.ppt;*.pptx";
            List<string> list = new List<string>();
            //ofd.DefaultExt = ".ppt|.pptx";
            // ofd.Filter = "ppt|*.ppt/*.pptx";
            if (ofd.ShowDialog() == true)
            {
                string path = ofd.FileName;
                if (Setting.Default.a==null)
                {
                    Setting.Default.a = new System.Collections.Specialized.StringCollection();
                }
                Setting.Default.a.Add(path);
                Setting.Default.Save();
                Process cmd = new Process();
                cmd.StartInfo.FileName = @"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE";
                cmd.StartInfo.Arguments =path;
                cmd.Start();
               // ofd.FileName

                //打开
            }
        }

        private void LstMain_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Console.WriteLine(LstMain.SelectedItem);
        }
    }
}
