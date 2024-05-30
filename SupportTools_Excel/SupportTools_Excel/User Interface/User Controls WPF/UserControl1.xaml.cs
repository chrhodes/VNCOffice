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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SupportTools_Excel.User_Interface
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var message = string.Format("Hi from UserControl1\nName: {0}, Uri:{1}\nNameDP: {2} UriDP:{3}", 
                tfsPicker.Name, tfsPicker.Uri,
                tfsPicker.NameDP, tfsPicker.UriDP);
            MessageBox.Show(message);
        }
    }
}
