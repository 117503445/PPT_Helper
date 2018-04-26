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

namespace PPT_Helper
{
    /// <summary>
    /// UMessageBox.xaml 的交互逻辑
    /// </summary>
    public partial class UMessageBox : UserControl
    {
        public event EventHandler ChooseOk;

        public UMessageBox()
        {
            InitializeComponent();
        }
        public UMessageBox(string message, EventHandler callBack)
        {
            Message = message;
            ChooseOk = callBack;
        }
        public string Message
        {
            get { return (string)GetValue(MessageProperty); }
            set { SetValue(MessageProperty, value); }
        }

        public static readonly DependencyProperty MessageProperty =
            DependencyProperty.Register("Message", typeof(string), typeof(UMessageBox), new PropertyMetadata("", new PropertyChangedCallback(Message_Changed)));

        void OnMessageChanged()
        {
            if (IsLoaded)
            {
                Tbk1.Text = Message;
            }
        }

        private static void Message_Changed(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            ((UMessageBox)d).OnMessageChanged();
        }

        private void Label2_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ChooseOk?.Invoke(this, new EventArgs());
        }
    }
}
