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
using System.Runtime.InteropServices;

namespace PPT_Helper
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class StartWindow : Window
    {
        List<string> list = new List<string>();
        /// <summary>
        /// 定义PowerPoint应用程序对象
        /// </summary>
        Microsoft.Office.Interop.PowerPoint.Application pptApplication;
        MainWindow MainWindow;
        public StartWindow()
        {
            InitializeComponent();
            try
            {
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as Microsoft.Office.Interop.PowerPoint.Application;
                MainWindow = new MainWindow();
                MainWindow.Show();
                this.Close();
            }
            catch
            {
                list = File.ReadAllLines(AppDomain.CurrentDomain.BaseDirectory + "path.txt", Encoding.Default).ToList();
                FlashLst();
            }
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "演示文稿|*.ppt;*.pptx",
                InitialDirectory = @"F:\文件\数学全品一轮复习\PPT听课手册配套课件"
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
            MainWindow = new MainWindow();
            MainWindow.Show();
            this.Close();
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
            MainWindow = new MainWindow();
            MainWindow.Show();
            this.Close();
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
            System.Threading.Thread.Sleep(5000);
        }
        private void FlashLst()
        {
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
