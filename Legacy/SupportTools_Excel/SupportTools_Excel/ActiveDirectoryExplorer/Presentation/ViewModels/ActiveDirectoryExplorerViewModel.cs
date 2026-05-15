using Prism.Commands;

using VNC;
using VNC.Core.Mvvm;

using SupportTools_Excel.Presentation.ModelWrappers;
using SupportTools_Excel.Presentation.Views;
using System;
using Prism.Events;
using SupportTools_Excel.Core.Presentation.ViewModels;
using SupportTools_Excel.ActiveDirectoryExplorer.Presentation.Views;
using System.DirectoryServices;

namespace SupportTools_Excel.ActiveDirectoryExplorer.Presentation.ViewModels
{
    public class ActiveDirectoryExplorerViewModel : ViewModelBase, IActiveDirectoryViewModel
    {
        #region Constructors and Load

        // View First
        // View creates new ViewModel in code or Xaml
        // or ViewModel passed into View constructor

        public ActiveDirectoryExplorerViewModel()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //ActiveDirectory = new ActiveDirectoryWrapper(new Domain.ActiveDirectory());

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // ViewModel First
        // Calling base(view) wires this ViewModel into the View

        public ActiveDirectoryExplorerViewModel(Views.ActiveDirectoryExplorer view) : base(view)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Initialize any controls and/or properties that need to be

            DoSomethingCommand = new DelegateCommand(OnDoSomethingExecute, OnDoSomethingCanExecute);


            Message_DoubleClick_Command = new DelegateCommand(Message_DoubleClick);

            //InitializeRows();

            // Put this in InitializeViewModel or Constructor

            FindUserCommand = new DelegateCommand(OnFindUserExecute, OnFindUserCanExecute);
            AddUserCommand = new DelegateCommand(OnAddUserExecute, OnAddUserCanExecute);

            SearchPattern_DoubleClick_Command = new DelegateCommand(SearchPattern_DoubleClick);

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Fields

        #endregion

        #region Properties



        string _message = "Click Button to do something";
        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                _message = value;
                OnPropertyChanged();
            }
        }

        private string _userName;
        public string UserName
        {
            get => _userName;
            set
            {
                if (_userName == value)
                    return;
                _userName = value;
                OnPropertyChanged();
            }
        }

        private string _aDEntryName;
        public string ADEntryName
        {
            get => _aDEntryName;
            set
            {
                if (_aDEntryName == value)
                    return;
                _aDEntryName = value;
                OnPropertyChanged();
            }
        }

        // Put these in ViewModel
        //public string SearchPattern { get; set; }
        private string _SearchPattern;
        //public string SearchPattern { get; set; }
        public string SearchPattern
        {
            get => _SearchPattern;
            set
            {
                if (_SearchPattern == value) return;
                _SearchPattern = value;
                OnPropertyChanged();
            }
        }

        public string SearchPatternToolTip { get; set; }

        // Put this in ViewModel Commands Region

        #region Command SearchPattern DoubleClick

        public DelegateCommand SearchPattern_DoubleClick_Command { get; set; }

        public void SearchPattern_DoubleClick()
        {
            Message = "SearchPattern_DoubleClick";
        }

        #endregion


        // TODO(crhodes)
        // This is for a Grid or List

        public System.Collections.ObjectModel.ObservableCollection<string> SelectedFruits { get; set; }

        // public System.Collections.ObjectModel.ObservableCollection<ActiveDirectoryWrapper> Rows { get; set; }

        // // and the SelectedItem in the Grid or List

        // ActiveDirectoryWrapper _selectedItem;
        // public ActiveDirectoryWrapper SelectedItem
        // {
        // get
        // {
        // return _selectedItem;
        // }
        // set
        // {
        // _selectedItem = value;
        // OnPropertyChanged();
        // }
        // }

        // Don't forget to uncomment InitializeRows in Constructors

        // void InitializeRows()
        // {
        // Rows = new System.Collections.ObjectModel.ObservableCollection<ActiveDirectoryWrapper>();
        // Rows.Add(new ActiveDirectoryWrapper(new Domain.ActiveDirectory(){ StringProperty ="Red", IntProperty = 1}));
        // Rows.Add(new ActiveDirectoryWrapper(new Domain.ActiveDirectory(){ StringProperty = "Green", IntProperty = 2 }));
        // Rows.Add(new ActiveDirectoryWrapper(new Domain.ActiveDirectory(){ StringProperty = "Blue", IntProperty = 3 }));

        // OnPropertyChanged("Rows");
        // }


        #endregion

        #region Commands

        #region DoSomething Command

        public DelegateCommand DoSomethingCommand { get; set; }
        public string DoSomethingContent { get; set; } = "Update Actions for selected shapes";
        public string DoSomethingToolTip { get; set; } = "ToolTip for DoSomething Button";

        public void OnDoSomethingExecute()
        {
            // TODO(crhodes)
            // Do something amazing.

            Message = "Cool, you did something!";
        }

        public bool OnDoSomethingCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.

            return true;
        }

        #endregion

        #region Control Commands (Not Buttons)

        public DelegateCommand Message_DoubleClick_Command { get; set; }

        public void Message_DoubleClick()
        {
            Message = "Message DoubleClicked!";
        }

        #endregion

        #region FindUser Command

        public DelegateCommand FindUserCommand { get; set; }
        public string FindUserContent { get; set; } = "FindUser";
        public string FindUserToolTip { get; set; } = "FindUser ToolTip";
        // Can get fancy and use Resources
        //public string FindUserContent { get; set; } = "ViewName_FindUserContent";
        //public string FindUserToolTip { get; set; } = "ViewName_FindUserContentToolTip";

        // Put these in Resource File

        //    <system:String x:Key="ViewName_FindUserContent">FindUser</system:String>
        //    <system:String x:Key="ViewName_FindUserContentToolTip">FindUser ToolTip</system:String>  

        public void OnFindUserExecute()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called FindUser";
            // Maybe put FindUser in Application Area
            Common.EventAggregator.GetEvent<FindUserEvent>().Publish();
            FindUser();
        }

        // Put this in places that listen for event
        //Common.EventAggregator.GetEvent<FindUserEvent>().Subscribe(FindUser);

        public bool OnFindUserCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        private void FindUser()
        {
            string userName = ((Views.ActiveDirectoryExplorer)View).teSearchPattern.Text;

            userName = SearchPattern;

            try
            {
                // create LDAP connection object  

                var name = ((Views.ActiveDirectoryExplorer)View).wucActiveDirectory_Picker.Name;
                var dnsHostName = ((Views.ActiveDirectoryExplorer)View).wucActiveDirectory_Picker.DNSHostName;
                var defaultNamingContext = ((Views.ActiveDirectoryExplorer)View).wucActiveDirectory_Picker.DefaultNamingContext;

                DirectoryEntry myLdapConnection = CreateDirectoryEntry(dnsHostName, defaultNamingContext);

                // create search object which operates on LDAP connection object  
                // and set search object to only find the user specified  

                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(cn=" + userName + ")";

                //List<String> requiredProperties = new List<string>();


                SearchResult result = search.FindOne();

                if (result != null)
                {
                    ((Views.ActiveDirectoryExplorer)View).lbeResults.Clear();

                    // user exists, cycle through LDAP fields (cn, telephonenumber etc.)  

                    ResultPropertyCollection fields = result.Properties;

                    foreach (String ldapField in fields.PropertyNames)
                    {
                        // cycle through objects in each field e.g. group membership  
                        // (for many fields there will only be one object such as name)  

                        foreach (Object myCollection in fields[ldapField])
                        {
                            string field = String.Format(">{0,30}< : >{1}<\n", ldapField.Trim(), myCollection.ToString());
                            ((Views.ActiveDirectoryExplorer)View).lbeResults.Text += field;
                        }
                    }
                }

                else
                {
                    Console.WriteLine("User not found!");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("Exception caught:\n\n" + ex.ToString());
            }
        }

        DirectoryEntry CreateDirectoryEntry(string dnsHostName, string defaultNamingContext)
        {
            // create and return new LDAP connection with desired settings  

            DirectoryEntry ldapConnection = new DirectoryEntry(dnsHostName);
            //DirectoryEntry ldapConnection = new DirectoryEntry();
            //ldapConnection.Path = "LDAP://OU=BDUsers,OU=AMERICAS,OU=CA154,DC=bdx,DC=com";
            ldapConnection.Path = string.Format("LDAP://{0}", defaultNamingContext);
            //ldapConnection.Path = @"LDAP://CN=Users,DC=VNC,DC=LOCAL";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

            return ldapConnection;
        }

        #endregion

        #region AddUser Command

        public DelegateCommand AddUserCommand { get; set; }
        public string AddUserContent { get; set; } = "AddUser";
        public string AddUserToolTip { get; set; } = "AddUser ToolTip";
        // Can get fancy and use Resources
        //public string AddUserContent { get; set; } = "ViewName_AddUserContent";
        //public string AddUserToolTip { get; set; } = "ViewName_AddUserContentToolTip";

        // Put these in Resource File

        //    <system:String x:Key="ViewName_AddUserContent">AddUser</system:String>
        //    <system:String x:Key="ViewName_AddUserContentToolTip">AddUser ToolTip</system:String>  

        public void OnAddUserExecute()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called AddUser";
            Common.EventAggregator.GetEvent<AddUserEvent>().Publish();
        }



        // Put this in places that listen for event
        //Common.EventAggregator.GetEvent<AddUserEvent>().Subscribe(AddUser);

        public bool OnAddUserCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        #endregion Commands

    }
}
