using System;
using System.Collections.Generic;
using System.Linq;
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
using System.Windows.Shapes;

namespace PPT_Helper
{
    /// <summary>
    /// InkEditWindow.xaml 的交互逻辑
    /// </summary>
    public partial class InkEditWindow : Window
    {
        public InkEditWindow()
        {
            InitializeComponent();

            drawingAttributes.Width = 5;
            drawingAttributes.Height = 5;
            drawingAttributes.Color = Colors.DeepSkyBlue;
            drawingAttributes.FitToCurve = true;
            
        }
        public DrawingAttributes drawingAttributes = new DrawingAttributes();
        private void RadioButton_Click(object sender, RoutedEventArgs e)
        {
            if (RdoBlack.IsChecked == true)
            {
                drawingAttributes.Color = Colors.Black;
            }
            else if (RdoBlue.IsChecked == true)
            {
                drawingAttributes.Color = Colors.DeepSkyBlue;
            }
            else if (RdoRed.IsChecked == true)
            {
                drawingAttributes.Color = Colors.Red;
            }
            else if (RdoWhite.IsChecked == true)
            {
                drawingAttributes.Color = Colors.White;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Hide();
            e.Cancel = true;
        }


        private void Slider_DragCompleted(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
        {
            drawingAttributes.Width = Slider.Value;
            drawingAttributes.Height = Slider.Value;
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            Top = SystemParameters.PrimaryScreenHeight - 300;
            Left =( SystemParameters.PrimaryScreenWidth - Width) / 2;
        }
    }
}
