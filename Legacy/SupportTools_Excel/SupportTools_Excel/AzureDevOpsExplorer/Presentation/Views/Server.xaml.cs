using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Git.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ViewModels;
using SupportTools_Excel.Core.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

using VNCTFS = VNC.TFS;
using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Presentation.Views
{
    public partial class Server : UserControl, IView
    {

        // TODO(crhodes)
        // Seems like these belong somewhere other than View (Server).
        // Should they go in ViewModel or a Domain Object.
        public static TfsConfigurationServer ConfigurationServer { get; set; }

        // These are updated when the Team Project Collection Changes

        public static IBuildServer BuildServer { get; set; }
        public static ICommonStructureService CommonStructureService { get; set; }
        public static GitRepositoryService GitRepositoryService { get; set; }
        public static IIdentityManagementService IdentityManagementService { get; set; }
        public static TestManagementService TestManagementService { get; set; }
        public static TfsTeamProjectCollection TfsTeamProjectCollection { get; set; }
        public static VersionControlServer VersionControlServer { get; set; }
        public static WorkItemStore WorkItemStore { get; set; }

        #region Constructors and Load

        // ViewModel First.  ViewModel creates View 
        // and sets DataContext by setting ViewModel property

        public Server()
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            // TODO(crhodes)
            // Until we call the other constructor, go get a ViewModel
            // Could also do this declaratively in Xaml.

            ViewModel = new ServerViewModel();

            InitializeView();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // View First.  View is passed ViewModel through Injection
        // or can declare ViewModel in Xaml

        public Server(IAZDOServerViewModel viewModel)
        {
            long startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            ViewModel = viewModel;

            InitializeView();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeView()
        {
            long startTicks = Log.VIEW("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Perform any initialization or configuration of View

            serverProvider.PopulateControlFromFile(Common.cCONFIG_FILE);

            liTeamProjectCollection.Visibility = Visibility.Hidden;
            liTeamProjectCollection2.Visibility = Visibility.Hidden;

            btnLoad_TFS_Collections.Visibility = Visibility.Hidden;

            //lgMain.IsCollapsed = true;

            Log.VIEW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Properties

        private IViewModel _viewModel;

        public IViewModel ViewModel
        {
            get { return _viewModel; }

            set
            {
                _viewModel = value;
                DataContext = _viewModel;
            }
        }

        #endregion 

        #region Event Handlers

        private void btnGetConfigurationServerInfo_Click(object sender, RoutedEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Common.EventAggregator.GetEvent<GetConfigurationServerInfoEvent>().Publish(serverProvider);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void serverProvider_Changed()
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            ServerChanged();

            btnLoad_TFS_Collections.Visibility = Visibility.Visible;

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        private void cbeTeamProjectCollections_SelectedIndexChanged(object sender, RoutedEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            try
            {
                var s1 = cbeTeamProjectCollections.Text;
                var s2 = cbeTeamProjectCollections.SelectedIndex;
                var s3 = cbeTeamProjectCollections.SelectedItemValue;
                var s4 = cbeTeamProjectCollections.SelectedItem;
                var s5 = cbeTeamProjectCollections.SelectedText;

                var s6 = ((ServerViewModel)ViewModel).SelectedItem;
                // The Server may have been changed - which clears the TeamProjectCollections.

                if (string.IsNullOrEmpty(cbeTeamProjectCollections.Text))
                {
                    Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
                    return;
                }

                Uri tpcUri = new Uri(cbeTeamProjectCollections.SelectedItem.ToString());

                // TODO(crhodes)
                // This might be the place to handle authentication.

                TfsTeamProjectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(tpcUri);

                if (TfsTeamProjectCollection == null)
                {
                    MessageBox.Show("Cannot GetTeamProjectCollection");
                    Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
                    return;
                }

                // Each time we change TeamProjectCollection, 
                // go get all the services we will use,
                // and put in _<member> variables.

                Populate_TPC_Services(TfsTeamProjectCollection);

                Common.EventAggregator.GetEvent<PopulateTeamProjectsEvent>().Publish();

                lgMain.IsCollapsed = true;

                // Need to handle this as an event or maybe property that wucTaskPane_TFS can bind to

                Common.EventAggregator.GetEvent<EnableMainUIEvent>().Publish(Visibility.Visible);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void btnLoad_TFS_Collections_Click(object sender, RoutedEventArgs e)
        {
            long startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            PopulateTeamProjectCollections();
            liTeamProjectCollection.Visibility = Visibility.Visible;
            liTeamProjectCollection2.Visibility = Visibility.Visible;

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ServerChanged()
        {
            long startTicks = Log.VIEW("Enter", Common.LOG_CATEGORY);

            try
            {
                if (ConfigurationServer != null)
                {
                    ConfigurationServer.Dispose();
                }

                // TODO(crhodes)
                // This might be the place to handle authentication.

                ConfigurationServer = VNCTFS.Helper.Get_ConfigurationServer(serverProvider.Uri);

                btnLoad_TFS_Collections.Visibility = Visibility.Visible;

                liTeamProjectCollection.Visibility = Visibility.Hidden;
                liTeamProjectCollection2.Visibility = Visibility.Hidden;

                Common.EventAggregator.GetEvent<EnableMainUIEvent>().Publish(Visibility.Hidden);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.VIEW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void PopulateTeamProjectCollections()
        {
            long startTicks = Log.VIEW("Enter", Common.LOG_CATEGORY);

            try
            {
                // Get the Team Project Collections

                ReadOnlyCollection<CatalogNode> projectCollectionNodes = VNCTFS.Helper.Get_TeamProjectCollectionNodes(ConfigurationServer);

                // Populate

                DevExpress.Xpf.Editors.ListItemCollection itemCol = cbeTeamProjectCollections.Items;

                itemCol.BeginUpdate();

                itemCol.Clear();

                foreach (CatalogNode teamProjectCollectionNode in projectCollectionNodes)
                {
                    TfsTeamProjectCollection teamProjectCollection = VNCTFS.Helper.Get_TeamProjectCollection(ConfigurationServer, teamProjectCollectionNode);

                    // TODO (crhodes):
                    // Maybe a class that contains a friendly name and a URI so populating Team Projects is easier.

                    itemCol.Add(teamProjectCollection.Uri);
                    //itemCol.Add(GetTeamProjectCollectionName(teamProjectCollection));
                    ((ServerViewModel)ViewModel).TeamProjectCollections.Add($"{teamProjectCollection.Uri}");
                    Log.Trace($"Added: {teamProjectCollection.Uri}", Common.LOG_CATEGORY, startTicks);
                }

                cbeTeamProjectCollections.SelectedIndex = -1;

                itemCol.EndUpdate();

                cbeTeamProjectCollections.ItemsSource = itemCol;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                // Should we throw?
            }

            Log.VIEW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Populate_TPC_Services(TfsTeamProjectCollection tpc)
        {
            long startTicks = Log.VIEW("Enter", Common.LOG_CATEGORY);

            try
            {

                BuildServer = null;
                CommonStructureService = null;
                GitRepositoryService = null;
                IdentityManagementService = null;
                TestManagementService = null;
                VersionControlServer = null;
                WorkItemStore = null;

                XlHlp.DisplayInWatchWindow("tpc.GetService<IBuildServer>()", startTicks);
                BuildServer = tpc.GetService<IBuildServer>();

                XlHlp.DisplayInWatchWindow("tpc.GetService<ICommonStructureService>()", startTicks);
                CommonStructureService = tpc.GetService<ICommonStructureService>();

                XlHlp.DisplayInWatchWindow("tpc.GetService<GitRepositoryService>()", startTicks);
                GitRepositoryService = tpc.GetService<GitRepositoryService>();

                XlHlp.DisplayInWatchWindow("tpc.GetService<IIdentityManagementService>()", startTicks);
                IdentityManagementService = tpc.GetService<IIdentityManagementService>();

                XlHlp.DisplayInWatchWindow("tpc.GetService<ITestManagementService>()", startTicks);
                TestManagementService = (TestManagementService)tpc.GetService<ITestManagementService>();

                XlHlp.DisplayInWatchWindow("VNCTFS.Helper.Get_VersionControlServer(tpc)", startTicks);
                VersionControlServer = VNCTFS.Helper.Get_VersionControlServer(tpc);

                XlHlp.DisplayInWatchWindow("new WorkItemStore(tpc)", startTicks);
                WorkItemStore = new WorkItemStore(tpc);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                // Should we throw?
            }

            Log.VIEW("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
