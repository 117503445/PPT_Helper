using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using PPt = Microsoft.Office.Interop.PowerPoint;
using Library;
namespace PPT_Helper
{
    enum DialogTask
    {
        None,
        Exit,
        Message,
        NullPPt
    }
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        DialogInventory<string> DialogInventory = new DialogInventory<string>();
        DialogTask DialogTask = DialogTask.None;
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
        PPt.Slide lastslide;
        /// <summary>
        /// 幻灯片的数量
        /// </summary>
        int slidesCount;
        /// <summary>
        /// 幻灯片的索引
        /// </summary>
        int slideIndex;
        /// <summary>
        /// 每页PPT都有的INK
        /// </summary>
        StrokeCollection[] strokes;
        InkEditWindow inkEditWindow = new InkEditWindow();
        internal Edit_Community.MultiInkCanvasWithTool MultiInkCanvasWithTool;
        private bool isHide = false;
        public bool IsHide
        {
            get => isHide;
            set
            {
                if (isHide != value)
                {
                    isHide = value;
                    ImgHide.IsChecked = !value;
                    if (value)
                    {
                        GridLeft.Visibility = Visibility.Hidden;
                        GridRight.Visibility = Visibility.Hidden;
                        GridExit.Visibility = Visibility.Hidden;
                        MultiInkCanvasWithTool.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        GridLeft.Visibility = Visibility.Visible;
                        GridRight.Visibility = Visibility.Visible;
                        GridExit.Visibility = Visibility.Visible;
                        MultiInkCanvasWithTool.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
        }
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            MultiInkCanvasWithTool = new Edit_Community.MultiInkCanvasWithTool()
            {
                IsTransparentStyle = false,
                InkMenuSelectIndex = 0
            };
            this.GridMain.Children.Add(MultiInkCanvasWithTool);
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
                    //lastslide = slide;
                    slideIndex = slide.SlideIndex;
                }
                catch
                {
                    // 在阅读模式下出现异常时，通过下面的方式来获得当前选中的幻灯片对象
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                    slideIndex = slide.SlideIndex;
                }
            }
            strokes = new StrokeCollection[slidesCount];
            for (int i = 0; i < strokes.Length; i++)
            {
                strokes[i] = new StrokeCollection();
            }
            SwitchInk();
        }
        private void SwitchInk()
        {
            MultiInkCanvasWithTool.Load(strokes[slide.SlideIndex - 1]);
            LblLeft.Content = slide.SlideIndex + "/" + slidesCount;
            LblRight.Content = slide.SlideIndex + "/" + slidesCount;
        }
        private void ImgPrev_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                slideIndex = slide.SlideIndex - 1;
            }
            catch
            {
                MessageBox.Show("PPt已关闭.", "Error");
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
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
                slideIndex = 1;
                if (sender.Equals(ImgPrevLeft))
                {
                    ShowMessageBox(new Point(80, 40), "已经是第一页了", DialogTask.Message, DialogX.Left, DialogY.Buttom);
                }
                else
                {
                    ShowMessageBox(new Point(220, 40), "已经是第一页了", DialogTask.Message, DialogX.Right, DialogY.Buttom);
                }
            }
            SwitchInk();
        }
        private void ImgNext_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                slideIndex = slide.SlideIndex + 1;
            }
            catch
            {
                MessageBox.Show("PPt已关闭.", "Error");
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            if (slideIndex > slidesCount)
            {
                if (sender.Equals(ImgNextLeft))
                {
                    ShowMessageBox(new Point(200, 40), "已经是最后一页了", DialogTask.Message, DialogX.Left, DialogY.Buttom);
                }
                else
                {
                    ShowMessageBox(new Point(80, 40), "已经是最后一页了", DialogTask.Message, DialogX.Right, DialogY.Buttom);
                }
                slideIndex = slidesCount;
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
            SwitchInk();
        }
        private void ImgExit_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ShowMessageBox(new Point(0, 80), "是否退出? 点[确定]立即退出.", DialogTask.Exit, DialogX.Right, DialogY.Buttom);
        }
        private void ImgHide_MouseUp(object sender, MouseButtonEventArgs e)
        {
            IsHide = !IsHide;
        }
        private void ShowMessageBox(Point point, string message, DialogTask dialogTask, DialogX dialogX,DialogY dialogY)
        {
            this.DialogTask = dialogTask;
            this.GridDialogBack.Visibility = Visibility.Visible;
            this.DialogInventory.Show("dialog", new DialogInfo( new UMessageBox(message, new EventHandler(MessageBox_MouseUp)), point, dialogX,dialogY,DialogType.Dialog, this.GridDialog));
        }
        private void MessageBox_MouseUp(object sender, EventArgs e)
        {
            if (DialogTask == DialogTask.Exit)
            {
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            HideMessageBox();
        }
        private void HideMessageBox()
        {
            this.DialogTask = DialogTask.None;
            this.GridDialogBack.Visibility = Visibility.Hidden;
            this.DialogInventory.Hide("dialog");
        }
        private void GridDialogBack_MouseUp(object sender, MouseButtonEventArgs e)
        {
            HideMessageBox();
        }
    }

}
