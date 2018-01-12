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
using System.Windows.Threading;

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
        InkCanvas[] inkCanvas;
        InkEditWindow inkEditWindow=new InkEditWindow();
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
                    //lastslide = slide;
                }
                catch
                {
                    // 在阅读模式下出现异常时，通过下面的方式来获得当前选中的幻灯片对象
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }

            }

            inkCanvas = new InkCanvas[slidesCount];
            for (int i = 0; i < inkCanvas.Length; i++)
            {
                inkCanvas[i] = new InkCanvas();
               
            }
            foreach (var item in inkCanvas)
            {
                item.Visibility = Visibility.Hidden;
                var bc = new BrushConverter();
                item.Background = (Brush)new BrushConverter().ConvertFrom("#02FFFFFF");
                grid.Children.Add(item);
                Grid.SetColumnSpan(item, 3);
                Grid.SetRowSpan(item, 2);
                item.DefaultDrawingAttributes = inkEditWindow.drawingAttributes;
                item.StrokeCollected += Item_StrokeCollected;
            }
                Panel.SetZIndex(StpTools, 1);//使工具栏置顶
            Panel.SetZIndex(StpRight, 1);//使工具栏置顶
            Panel.SetZIndex(StpLeft, 1);//使工具栏置顶
        }

        private void Item_StrokeCollected(object sender, InkCanvasStrokeCollectedEventArgs e)
        {
            inkEditWindow.Hide();
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
            SwitchInk();
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
         SwitchInk();
        }

        private void SwitchInk()
        {
            foreach (var item in inkCanvas)
            {
                item.Visibility = Visibility.Hidden;
            }
            inkCanvas[slide.SlideIndex - 1].Visibility = Visibility.Visible;
        }


        private void ImgInk_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (inkMode)
            {
                if (inkEditWindow.IsVisible)
                {
                    inkEditWindow.Hide();
                }
                else
                {
                    inkEditWindow.Show();
                    inkEditWindow.Topmost = true;
                }
            }
            foreach (var item in inkCanvas)
            {
                item.EditingMode = InkCanvasEditingMode.Ink;
            }
            eraserMode = false;
            inkMode = true;
            ImgMouse.Source = new BitmapImage(new Uri("/Resources/Tools/UnChecked/Mouse.jpg", UriKind.RelativeOrAbsolute));
            ImgInk.Source = new BitmapImage(new Uri("/Resources/Tools/Checked/Pen.jpg", UriKind.RelativeOrAbsolute));
            ImgEraser.Source= new BitmapImage(new Uri("/Resources/Tools/Unchecked/Eraser.jpg", UriKind.RelativeOrAbsolute));
            SwitchInk();

        }

        private void ImgMouse_MouseDown(object sender, MouseButtonEventArgs e)
        {
            eraserMode = false;
            inkMode = false;
            ImgMouse.Source = new BitmapImage(new Uri("/Resources/Tools/Checked/Mouse.jpg", UriKind.RelativeOrAbsolute));
            ImgInk.Source = new BitmapImage(new Uri("/Resources/Tools/UnChecked/Pen.jpg", UriKind.RelativeOrAbsolute));
            ImgEraser.Source = new BitmapImage(new Uri("/Resources/Tools/Unchecked/Eraser.jpg", UriKind.RelativeOrAbsolute));
            foreach (var item in inkCanvas)
            {
                item.Visibility = Visibility.Hidden;
            }

        }
        /// <summary>
        /// 是否在橡皮模式
        /// </summary>
        bool eraserMode = false;
        bool inkMode = false;
        private void ImgEraser_MouseDown(object sender, MouseButtonEventArgs e)
        {
            foreach (var item in inkCanvas)
            {
                item.EditingMode = InkCanvasEditingMode.EraseByStroke;
            }
            if (eraserMode)
            {
                if (MessageBox.Show("是否全部清理", "", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    inkCanvas[slide.SlideIndex - 1].Strokes.Clear();
                }
            }
            else
            {
                ImgMouse.Source = new BitmapImage(new Uri("/Resources/Tools/UnChecked/Mouse.jpg", UriKind.RelativeOrAbsolute));
                ImgInk.Source = new BitmapImage(new Uri("/Resources/Tools/UnChecked/Pen.jpg", UriKind.RelativeOrAbsolute));
                ImgEraser.Source = new BitmapImage(new Uri("/Resources/Tools/checked/Eraser.jpg", UriKind.RelativeOrAbsolute));
                eraserMode = true;
                inkMode = false;
            }

        }
        private void ImgClose_MouseDown(object sender, MouseButtonEventArgs e)
        {
#if !DEBUG
            if (MessageBox.Show("是否退出?", "退出", MessageBoxButton.YesNo) == MessageBoxResult.No)
            {
                return;
            }
#endif
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }

        private void ImgSetting_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (inkEditWindow.IsVisible)
            {
                inkEditWindow.Hide();
            }
            else
            {
                inkEditWindow.Show();
                inkEditWindow.Topmost = true;
            }

        }


    }
}
