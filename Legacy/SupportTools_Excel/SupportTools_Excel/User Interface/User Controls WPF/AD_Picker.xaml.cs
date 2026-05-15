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

namespace SupportTools_Excel.User_Interface.User_Controls_WPF
{
    public partial class AD_Picker : UserControl
    {

        #region Constructors and Load
        
        public AD_Picker()
        {
            InitializeComponent();
        }
        #endregion

        #region Properties and Fields

        // TODO:
        //  Add properties, backing fields, and dependency properties to match the attributes from each XML element
        //  Convention is _<Attribute Name>, <AttributeName>, <AttributeName>DP
        //  where <AttributeName> is the name of the attrbute for the <Element>

        private string _Name;


        public string Name
        {
            get { return _Name; }
            set
            {
                _Name = value;
            }
        }

        public string NameDP
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(NameDPProperty);
            }
            set
            {
                SetValue(NameDPProperty, value);
            }
        }

        // TODO:
        //  Update the parameters to .Register(...) as needed

        public static readonly DependencyProperty NameDPProperty = DependencyProperty.Register(
            "NameDP", typeof(string), typeof(AD_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnNameDPChanged)));

        public string DefaultNamingContext
        {
            get { return _DefaultNamingContext; }
            set
            {
                _DefaultNamingContext = value;
            }
        }

        private string _DefaultNamingContext;
        public string DefaultNamingContextDP
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(DefaultNamingContextDPProperty);
            }
            set
            {
                SetValue(DefaultNamingContextDPProperty, value);
            }
        }
        // TODO:
        //  Update the parameters to .Register(...) as needed

        public static readonly DependencyProperty DefaultNamingContextDPProperty = DependencyProperty.Register(
            "DefaultNamingContextDP", typeof(string), typeof(AD_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnDefaultNamingContextDPChanged)));

        private string _DNSHostName;

        public string DNSHostName
        {
            get { return _DNSHostName; }
            set
            {
                _DNSHostName = value;
            }
        }

        public string DNSHostNameDP
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(DNSHostNameDPProperty);
            }
            set
            {
                SetValue(DNSHostNameDPProperty, value);
            }
        }

        // TODO:
        //  Update the parameters to .Register(...) as needed

        public static readonly DependencyProperty DNSHostNameDPProperty = DependencyProperty.Register(
            "DNSHostNameDP", typeof(string), typeof(AD_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnDNSHostNameDPChanged)));

        #endregion
        
        #region Delegates and Events
        
        public event ControlEvent ControlChanged;

        public delegate void ControlEvent();

        #endregion

        #region Event Handlers

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var v = sender;
            var item = ((System.Windows.Controls.ComboBox)v).SelectedItem;

            if (null == item)
            {
                // May have just opened new file and no item has been selected.
            	return;
            }

            _Name = ((System.Xml.XmlElement)item).Attributes["Name"].Value;
            NameDP = ((System.Xml.XmlElement)item).Attributes["Name"].Value;

            _DNSHostName = ((System.Xml.XmlElement)item).Attributes["DNSHostName"].Value;
            DNSHostNameDP = ((System.Xml.XmlElement)item).Attributes["DNSHostName"].Value;

            _DefaultNamingContext = ((System.Xml.XmlElement)item).Attributes["DefaultNamingContext"].Value;
            DefaultNamingContextDP = ((System.Xml.XmlElement)item).Attributes["DefaultNamingContext"].Value;

            ControlEvent fireEvent = Interlocked.CompareExchange(ref ControlChanged, null, null);

            if (null != fireEvent)
            {
            	fireEvent();
            }
        }

        protected virtual void OnNameDPChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnNameDPChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            AD_Picker picker = o as AD_Picker;
            if (picker != null)
                picker.OnNameDPChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnDNSHostNameDPChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnDNSHostNameDPChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            AD_Picker picker = o as AD_Picker;
            if (picker != null)
                picker.OnDNSHostNameDPChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnDefaultNamingContextDPChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnDefaultNamingContextDPChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            AD_Picker picker = o as AD_Picker;
            if (picker != null)
                picker.OnDefaultNamingContextDPChanged((string)e.OldValue, (string)e.NewValue);
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
