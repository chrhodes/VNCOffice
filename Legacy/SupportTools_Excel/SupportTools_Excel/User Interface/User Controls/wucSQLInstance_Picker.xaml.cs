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
using System.Xml;

using Microsoft.Win32;


namespace SupportTools_Excel.User_Interface.User_Controls
{
    public partial class wucSQLInstance_Picker : UserControl
    {

        #region Constructors and Load

        public wucSQLInstance_Picker()
        {
            InitializeComponent();
        }

        #endregion

        #region Properties and Fields

        // TODO:
        //  Add properties, backing fields, and dependency properties to match the attributes from each XML element
        //  Convention is _<Attribute Name>, <AttributeName>, <AttributeName>DP
        //  where <AttributeName> is the name of the attribute for the <Element>

        public List<string> Databases
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (List<string>)GetValue(DatabasesProperty);
            }
            set
            {
                SetValue(DatabasesProperty, value);
            }
        }

        public string ExpandedName
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(ExpandedNameProperty);
            }
            set
            {
                SetValue(ExpandedNameProperty, value);
            }
        }

        public string FullName
         {
         	// IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
         	get
         	{
         		return (string)GetValue(FullNameProperty);
         	}
         	set
         	{
         		SetValue(FullNameProperty, value);
         	}
         }


        public string IPv4Address
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, 
            // do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(IPv4AddressProperty);
            }
            set
            {
                SetValue(IPv4AddressProperty, value);
            }
        }

        public string Instance
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, 
            // do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(InstanceProperty);
            }
            set
            {
                SetValue(InstanceProperty, value);
            }
        }

        public string Port
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, 
            // do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(PortProperty);
            }
            set
            {
                SetValue(PortProperty, value);
            }
        }

        public string Server
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, 
            // do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(ServerProperty);
            }
            set
            {
                SetValue(ServerProperty, value);
            }
        }

        // TODO:
        //  Update the parameters to .Register(...) as needed

        public static readonly DependencyProperty DatabasesProperty = DependencyProperty.Register(
            "Databases", typeof(List<string>), typeof(wucSQLInstance_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnDatabasesChanged)));

        public static readonly DependencyProperty ExpandedNameProperty = DependencyProperty.Register(
            "ExpandedName", typeof(string), typeof(wucSQLInstance_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnExpandedNameChanged)));

        public static readonly DependencyProperty FullNameProperty = DependencyProperty.Register(
            "FullName", typeof(string), typeof(wucSQLInstance_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnFullNameChanged)));

        public static readonly DependencyProperty IPv4AddressProperty = DependencyProperty.Register(
            "IPv4Address", typeof(string), typeof(wucSQLInstance_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnIPv4AddressChanged)));

        public static readonly DependencyProperty InstanceProperty = DependencyProperty.Register(
            "Instance", typeof(string), typeof(wucSQLInstance_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnInstanceChanged)));

        public static readonly DependencyProperty PortProperty = DependencyProperty.Register(
            "Port", typeof(string), typeof(wucSQLInstance_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnPortChanged)));

        public static readonly DependencyProperty ServerProperty = DependencyProperty.Register(
            "Server", typeof(string), typeof(wucSQLInstance_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnServerChanged)));
        
        #endregion

        #region Delegates and Events

        public delegate void ControlEvent();
        public event ControlEvent ControlChanged;

        #endregion

        #region Event Handlers

         private static void OnExpandedNameChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
        	wucSQLInstance_Picker wucSQLInstance_Picker = o as wucSQLInstance_Picker;
        	if (wucSQLInstance_Picker != null)
        		wucSQLInstance_Picker.OnExpandedNameChanged((string)e.OldValue, (string)e.NewValue);
        }
        
        protected virtual void OnExpandedNameChanged(string oldValue, string newValue)
        {
        	// TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnFullNameChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
         {
         	wucSQLInstance_Picker wucSQLInstance_Picker = o as wucSQLInstance_Picker;
         	if (wucSQLInstance_Picker != null)
         		wucSQLInstance_Picker.OnFullNameChanged((string)e.OldValue, (string)e.NewValue);
         }
         
         protected virtual void OnFullNameChanged(string oldValue, string newValue)
         {
         	// TODO: Add your property changed side-effects. Descendants can override as well.
         }

        private static void OnDatabasesChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            wucSQLInstance_Picker wucSQLInstance_Picker = o as wucSQLInstance_Picker;
            if (wucSQLInstance_Picker != null)
                wucSQLInstance_Picker.OnDatabasesChanged((List<string>)e.OldValue, (List<string>)e.NewValue);
        }

        protected virtual void OnDatabasesChanged(List<string> oldValue, List<string> newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private string GetFullName()
        {
            string updatedName;

            // Prefer HostName over IPs and InstanceNames over Port#s

            if (Instance.Length > 0 || Port.Length > 0)  // Have Instance Info           
            {
                updatedName = string.Format(@"{0}{1}",
                    Server.Length > 0 ? Server : IPv4Address,
                    Instance.Length > 0 ? @"\" + Instance : "," + Port
                    );
            }
            else
            {
                updatedName = Server.Length > 0 ? Server : IPv4Address;
            }

            return updatedName;
        }

        private string GetExpandedName()
        {
            string updatedName;

            if (Instance.Length > 0 || Port.Length > 0)
            {
                updatedName = string.Format(@"{0} ({1}) \ {2} ({3})",
                    Server.Length > 0 ? Server : "??",
                    IPv4Address.Length > 0 ? IPv4Address : "??",
                    Instance.Length > 0 ? Instance : "??",
                    Port.Length > 0 ? Port : "??");
            }
            else
            {
                updatedName = string.Format(@"{0} ({1}) \ <Default>",
                    Server.Length > 0 ? Server : "??",
                    IPv4Address.Length > 0 ? IPv4Address : "??");
            }

            return updatedName;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var v = sender;
            var item = ((System.Windows.Controls.ComboBox)v).SelectedItem;

            if (null == item)
            {
                // May have just opened new file and no item has been selected.
                return;
            }

            Server = ((System.Xml.XmlElement)item).Attributes["Server"].Value;
            IPv4Address = ((System.Xml.XmlElement)item).Attributes["IPv4Address"].Value;
            Instance = ((System.Xml.XmlElement)item).Attributes["Instance"].Value;
            Port = ((System.Xml.XmlElement)item).Attributes["Port"].Value;
            FullName = GetFullName();
            ExpandedName = GetExpandedName();
            Databases = GetDatabases((System.Xml.XmlElement)item);
 

            //Shouldn't have to do this in code.  Need to figure out markup.

            teExpandedName.Text = ExpandedName;
            teFullName.Text = FullName;

            ControlEvent fireEvent = Interlocked.CompareExchange(ref ControlChanged, null, null);

            if (null != fireEvent)
            {
                fireEvent();
            }
        }

        private List<string> GetDatabases(System.Xml.XmlElement instanceElement)
        {
            XmlNodeList databaseNodes = instanceElement.GetElementsByTagName("Database");

            List<string> databases = null; 
            //string[] database2 = new string[] { "One", "Two" };

            if (databaseNodes.Count > 0)
            {
                databases = new List<string>();

                foreach (XmlNode item in databaseNodes)
                {
                    databases.Add(item.Attributes["Name"].Value);
                }
            }

            return databases;
        }

        private static void OnPortChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
        	wucSQLInstance_Picker wucSQLServer_Picker = o as wucSQLInstance_Picker;
        	if (wucSQLServer_Picker != null)
        		wucSQLServer_Picker.OnPortChanged((string)e.OldValue, (string)e.NewValue);
        }
        
        protected virtual void OnPortChanged(string oldValue, string newValue)
        {
        	// TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnInstanceChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
        	wucSQLInstance_Picker wucSQLServer_Picker = o as wucSQLInstance_Picker;
        	if (wucSQLServer_Picker != null)
        		wucSQLServer_Picker.OnInstanceChanged((string)e.OldValue, (string)e.NewValue);
        }
        
        protected virtual void OnInstanceChanged(string oldValue, string newValue)
        {
        	// TODO: Add your property changed side-effects. Descendants can override as well.
        }

        protected virtual void OnIPv4AddressChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnIPv4AddressChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            wucSQLInstance_Picker picker = o as wucSQLInstance_Picker;
            if (picker != null)
                picker.OnIPv4AddressChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnServerChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnServerChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            wucSQLInstance_Picker picker = o as wucSQLInstance_Picker;
            if (picker != null)
                picker.OnServerChanged((string)e.OldValue, (string)e.NewValue);
        }

        private void Reload_Click(object sender, RoutedEventArgs e)
        {
            LoadDataFromFile();
        }


        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            //this.DataContext = this;
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
            openFileDialog1.InitialDirectory = @"C:\temp";

            if (true == openFileDialog1.ShowDialog())
            {
                string fileName = openFileDialog1.FileName;

                PopulateControlFromFile(fileName);
            }
        }


        #endregion
    }
}
