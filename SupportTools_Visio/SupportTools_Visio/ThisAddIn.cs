using System;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;

using DevExpress.Xpf.Core;

using VNC;

namespace SupportTools_Visio
{
    public partial class ThisAddIn
    {
        private static System.Windows.Application _XamlApp;

        private static Prism.Unity.PrismApplication _prismApplication;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {  
            Int64 startTicks = Log.APPLICATION_START("Enter", Common.LOG_CATEGORY);

            try
            {
                Globals.Ribbons.Ribbon.chkDisplayEvents.Checked = Common.DisplayEvents;
                Globals.Ribbons.Ribbon.chkEnableAppEvents.Checked = Common.HasAppEvents;

                if (Common.HasAppEvents)
                {
                    Common.AppEvents = new Events.VisioAppEvents();
                    Common.AppEvents.VisioApplication = Globals.ThisAddIn.Application;
                }

                Common.VisioHelper.VisioApplication = Globals.ThisAddIn.Application;

                // For the WPF stuff we do
                InitializeWPFApplication();
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
                throw (ex);
            }

            Log.APPLICATION_START("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Int64 startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);

            try
            {
                if (Common.HasAppEvents)
                {
                    Common.AppEvents = null;
                }

                UnLoadXamlApplicationResources();
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
                throw (ex);
            }

            Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
        }

        void InitializeWPFApplication()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Common.CurrentUser = new WindowsPrincipal(WindowsIdentity.GetCurrent());

            CreateXamlApplication();

            try
            {
                if (Data.Config.ADBypass)
                {
                    Common.IsAdministrator = true;
                    Common.IsBetaUser = true;
                    Common.IsDeveloper = true;
                }
                //else
                //{
                //    if (!Data.Config.AD_Users_AllowAll)
                //    {
                //        bool isAuthorizedUser = ADHelper.ADHelper.CheckGroupMembership(
                //            //"maward", 
                //            Common.CurrentUser.Identity.Name,
                //            SQLInformation.Data.Config.ADGroup_Users,
                //            SQLInformation.Data.Config.AD_Domain);

                //        if (!isAuthorizedUser)
                //        {
                //            MessageBox.Show(string.Format("You must be a member of {0}\\{1} to run this application.",
                //                SQLInformation.Data.Config.AD_Domain, SQLInformation.Data.Config.ADGroup_Users));
                //            return;
                //        }
                //    }

                //    Common.IsAdministrator = ADHelper.ADHelper.CheckDirectGroupMembership(
                //        Common.CurrentUser.Identity.Name,
                //        SQLInformation.Data.Config.ADGroup_Administrators,
                //        SQLInformation.Data.Config.AD_Domain);


                //    Common.IsBetaUser = ADHelper.ADHelper.CheckDirectGroupMembership(
                //        Common.CurrentUser.Identity.Name,
                //        SQLInformation.Data.Config.ADGroup_BetaUsers,
                //        SQLInformation.Data.Config.AD_Domain);

                //    Common.IsDeveloper = Common.CurrentUser.Identity.Name.Contains("crhodes") ? true : false;

                //    // Next lines are for testing UI only.  Comment out for normal operation.
                //    //Common.IsAdministrator = false;   
                //    //Common.IsBetaUser = false; 
                //    //Common.IsDeveloper = false;
                //}

                // Cannot do here as the Common.ApplicationDataSet has not been loaded.  Need to move here or do later.
                // For now this is in DXRibbonWindowMain();

                //var eventMessage = "Started";
                //SQLInformation.Helper.IndicateApplicationUsage(LOG_APPNAME, DateTime.Now, currentUser.Identity.Name, eventMessage);

                // Launch the main window.

                // Done from Ribbon

                //User_Interface.Windows.SplashScreen _window1 = new User_Interface.Windows.SplashScreen();
                //User_Interface.Windows.DXRibbonWindowMain _window1 = new User_Interface.Windows.DXRibbonWindowMain();

                //String windowArgs = string.Empty;
                // Check for arguments; if there are some build the path to the package out of the args.
                //if (args.Args.Length > 0 && args.Args[0] != null)
                //{
                //    for (int i = 0; i < args.Args.Length; ++i)
                //    {
                //        windowArgs = args.Args[i];
                //        switch (i)
                //        {
                //            case 0: // Patient Id
                //                //patientId = windowArgs;
                //                break;
                //        }
                //    }
                //}

                //_window1.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                MessageBox.Show(ex.InnerException.ToString());
            }

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        /// <summary>
        /// LoadXamlApplicationResources
        ///
        /// Creates Xaml Resources collection in System.Windows.Application
        /// for use in Hosted applications without App.Xaml
        /// </summary>

        private static void CreateXamlApplication()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Common.DeveloperMode = true;
            Common.WriteToDebugWindow("CreateXamlApplication()");
            Common.DeveloperMode = false;

            try
            {
                // TODO(crhodes)

                // Can we just create a PrismApplication?
                // Create a WPF Application
                _XamlApp = new System.Windows.Application();

                //_prismApplication = new Application.PrismApp();

                var defaultThemes = Theme.Themes;
                ApplicationThemeHelper.ApplicationThemeName = "MetropolisDark";

                // Load the resources

                // This works

                //var resources = System.Windows.Application.LoadComponent(
                //    new Uri("SupportTools_Excel;component/Resources/Xaml/Brushes.xaml", UriKind.Relative)) as System.Windows.ResourceDictionary;

                // Now lets try with

                var resources = System.Windows.Application.LoadComponent(
                    new Uri("SupportTools_Visio;component/Resources/Xaml/Application.xaml", UriKind.Relative)) as System.Windows.ResourceDictionary;

                //var resources = System.Windows.Application.LoadComponent(
                //    new Uri("pack:/SupportTools_Excel;:,,/Resources/Xaml/Application.xaml")) as System.Windows.ResourceDictionary;

                // Merge it on application level

                _XamlApp.Resources.MergedDictionaries.Add(resources);

                //_prismApplication.Resources.MergedDictionaries.Add(resources);
            }
            catch (Exception ex)
            {
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString());
                Common.DeveloperMode = false;
            }

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void UnLoadXamlApplicationResources()
        {
            Int64 startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);

            try
            {
                if (null != _XamlApp)
                {
                    _XamlApp.Shutdown();
                    _XamlApp = null;
                }
                if (null != _prismApplication)
                {
                    _prismApplication.Shutdown();
                    _prismApplication = null;
                }
            }
            catch (Exception ex)
            {
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString());
                Common.DeveloperMode = false;
            }

            Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

    }
}
