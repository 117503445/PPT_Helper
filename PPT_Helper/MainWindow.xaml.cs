using System;
using System.Collections.Generic;
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
using System.Runtime.InteropServices;
using PPt = Microsoft.Office.Interop.PowerPoint;
namespace PPT_Helper
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        /// <summary>
        /// 定义PowerPoint应用程序对象
        /// </summary>
        PPt.Application pptApplication;
        /// <summary>
        /// 定义演示文稿对象
        /// </summary>
        PPt.Presentation presentation;
        /// <summary>
        /// 定义幻灯片集合对象
        /// </summary>
        PPt.Slides slides;
        /// <summary>
        /// 定义单个幻灯片对象
        /// </summary>
        PPt.Slide slide;
        /// <summary>
        /// 幻灯片的数量
        /// </summary>
        int slidesCount;
        /// <summary>
        /// 幻灯片的索引
        /// </summary>
        int slideIndex;

        public MainWindow()
        {
            InitializeComponent();
            try
            {
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
            }
            catch
            {
                MessageBox.Show("请先启动遥控的幻灯片,此程序即将关闭", "Error");
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            if (pptApplication != null)
            {
                //获得演示文稿对象
                presentation = pptApplication.ActivePresentation;
                // 获得幻灯片对象集合
                slides = presentation.Slides;
                // 获得幻灯片的数量
                slidesCount = slides.Count;
                // 获得当前选中的幻灯片
                try
                {
                    // 在普通视图下这种方式可以获得当前选中的幻灯片对象
                    // 然而在阅读模式下，这种方式会出现异常
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // 在阅读模式下出现异常时，通过下面的方式来获得当前选中的幻灯片对象
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }
        private void BtnPrev_Click(object sender, RoutedEventArgs e)
        {
            slideIndex = slide.SlideIndex - 1;
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    // 在阅读模式下使用下面的方式来切换到上一张幻灯片
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                MessageBox.Show("已经是第一页了");
            }

        }
        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            slideIndex = slide.SlideIndex + 1;
            if (slideIndex > slidesCount)
            {
                MessageBox.Show("已经是最后一页了");
            }
            else
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    // 在阅读模式下使用下面的方式来切换到下一张幻灯片
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }
    }
}
