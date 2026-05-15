using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Win32;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    public partial class wucAssembly_Picker : UserControl
    {

        #region Constructors and Load

        public wucAssembly_Picker()
        {
            InitializeComponent();
        }

        #endregion

        #region Properties and Fields

        // TODO:
        //  Add properties, backing fields, and dependency properties to match the attributes from each XML element
        //  Convention is _<Attribute Name>, <AttributeName>, <AttributeName>DP
        //  where <AttributeName> is the name of the attribute for the <Element>

        private string _FullPath;

        public string FullPath
        {
            get { return _FullPath; }
            set
            {
                _FullPath = value;
            }
        }

        public string FullPathDP
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(FullPathDPProperty);
            }
            set
            {
                SetValue(FullPathDPProperty, value);
            }
        }

        // TODO:
        //  Update the parameters to .Register(...) as needed

        public static readonly DependencyProperty FullPathDPProperty = DependencyProperty.Register(
            "FullPathDP", typeof(string), typeof(wucAssembly_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnFullPathDPChanged)));

        #endregion

        #region Delegates and Events

        public delegate void ControlEvent();
        public event ControlEvent ControlChanged;

        #endregion

        #region Event Handlers

        private void ComboBoxEdit_SelectedIndexChanged(object sender, RoutedEventArgs e)
        {
            var v = sender;
            var item = ((DevExpress.Xpf.Editors.ComboBoxEdit)v).SelectedItem;

            if (null == item)
            {
                // May have just opened new file and no item has been selected.
                return;
            }

            _FullPath = ((System.Xml.XmlElement)item).Attributes["FullPath"].Value;
            FullPathDP = ((System.Xml.XmlElement)item).Attributes["FullPath"].Value;

            ControlEvent fireEvent = Interlocked.CompareExchange(ref ControlChanged, null, null);

            if (null != fireEvent)
            {
                fireEvent();
            }
        }

        private void ComboBoxEdit_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var v = sender;
            var item = ((System.Windows.Controls.ComboBox)v).SelectedItem;

            if (null == item)
            {
                // May have just opened new file and no item has been selected.
                return;
            }

            _FullPath = ((System.Xml.XmlElement)item).Attributes["FullPath"].Value;
            FullPathDP = ((System.Xml.XmlElement)item).Attributes["FullPath"].Value;

            ControlEvent fireEvent = Interlocked.CompareExchange(ref ControlChanged, null, null);

            if (null != fireEvent)
            {
                fireEvent();
            }
        }

        protected virtual void OnFullPathDPChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnFullPathDPChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            wucAssembly_Picker picker = o as wucAssembly_Picker;
            if (picker != null)
                picker.OnFullPathDPChanged((string)e.OldValue, (string)e.NewValue);
        }

        private void Reload_Click(object sender, RoutedEventArgs e)
        {
            LoadDataFromFile();
        }

        #endregion

        #region Main Methods

        public void PopulateControlFromFile(string fileNameAndPath)
        {
            comboBox.Source = new Uri(fileNameAndPath);
        }

        #endregion

        #region Private Methods

        private void LoadDataFromFile()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog1.FileName = "";

            if (true == openFileDialog1.ShowDialog())
            {
                string fileName = openFileDialog1.FileName;

                PopulateControlFromFile(fileName);
            }
        }


        #endregion
    }
}
