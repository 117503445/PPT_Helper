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
using System.IO;
namespace WpfApp1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        List<string> list = new List<string>();
        public MainWindow()
        {
            InitializeComponent();
            if (true)
            {
                //
            }
            list = File.ReadAllLines(AppDomain.CurrentDomain.BaseDirectory + "path.txt", Encoding.Default).ToList();
            FlashLst();
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "演示文稿|*.ppt;*.pptx",
                InitialDirectory = @"H:\Project\PPT_Helper\WpfApp1\bin\Debug"
            };

            if (ofd.ShowDialog() == true)
            {
                string path = ofd.FileName;
                if (!list.Contains(path))
                {
                    list.Add(path);
                }
                else
                {
                    int j = list.FindIndex(i => i.Equals(path));
                    string temp = list[j];
                    list[j] = list[0];
                    list[0] = temp;
                }

                OpenPPT(path);
            }
            FlashLst();
            SaveList();
        }

        private void LstMain_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            string path = LstMain.SelectedItem.ToString();
            int j = list.FindIndex(i => i.Equals(path));
            string temp = list[j];
            list[j] = list[0];
            list[0] = temp;
            FlashLst();
            OpenPPT(path);
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            list.Remove(LstMain.SelectedItem.ToString());
            FlashLst();
            SaveList();
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            list.Clear();
            FlashLst();
            SaveList();
        }
        private void OpenPPT(string path, string PPT_exe_Path = @"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE")
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = PPT_exe_Path;
            cmd.StartInfo.Arguments = path;
            cmd.Start();

        }
        private void FlashLst() {
            LstMain.Items.Clear();
            foreach (var item in list)
            {
                LstMain.Items.Add(item);
            }
        }
        private void SaveList()
        {
            File.WriteAllLines(AppDomain.CurrentDomain.BaseDirectory + "path.txt", list, Encoding.Default);
        }
    }
}
