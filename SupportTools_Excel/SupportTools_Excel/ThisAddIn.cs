using System;
using System.Security.Principal;
using System.Windows;

using DevExpress.Xpf.Core;

using VNC;

using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel
{
    public partial class ThisAddIn
    {
        // HACK(crhodes)
        // Have move away from TaskPanes

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneExcelUtil;

        //private static System.Windows.Application _XamlApp;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            long startTicks = Log.APPLICATION_START("Enter", Common.LOG_CATEGORY);

            try
            {
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow("ThisAddIn_Startup()");
                Common.DeveloperMode = false;

                Globals.Ribbons.Ribbon.chkDisplayEvents.Checked = Common.DisplayEvents;
                Globals.Ribbons.Ribbon.chkEnableAppEvents.Checked = Common.HasAppEvents;

                if (Common.HasAppEvents)
                {
                    Common.AppEvents = new Events.ExcelAppEvents();
                    Common.AppEvents.ExcelApplication = Globals.ThisAddIn.Application;
                }

                XlHlp.ExcelApplication = Globals.ThisAddIn.Application;

                // For the WPF stuff we do

                InitializeWPFApplication();
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);

                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString());
                Common.DeveloperMode = false;

                throw (ex);
            }

            Log.APPLICATION_START("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            long startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);

            try
            {
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow("ThisAddIn_Shutdown()");
                Common.DeveloperMode = false;

                if (Common.HasAppEvents)
                {
                    Common.AppEvents = null;
                }

                UnLoadXamlApplication();


            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);

                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString());
                Common.DeveloperMode = false;

                throw (ex);
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
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

        /// <summary>
        /// LoadXamlApplicationResources
        ///
        /// Creates Xaml Resources collection in System.Windows.Application
        /// for use in Hosted applications without App.Xaml
        /// </summary>

        void InitializeWPFApplication()
        {
            long startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

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
        ///

        private static void CreateXamlApplication()
        {
            long startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Common.DeveloperMode = true;
            Common.WriteToDebugWindow("CreateXamlApplication()");
            Common.DeveloperMode = false;

            try
            {
                // TODO(crhodes)
                // Can we just create a PrismApplication?

                // Create a WPF Application
                Common.XamlApplication = new System.Windows.Application();

                var defaultThemes = Theme.Themes;
                ApplicationThemeHelper.ApplicationThemeName = "MetropolisDark";

                // Load the resources

                // This works

                //var resources = System.Windows.Application.LoadComponent(
                //    new Uri("SupportTools_Excel;component/Resources/Xaml/Brushes.xaml", UriKind.Relative)) as System.Windows.ResourceDictionary;

                // Now lets try with

                var resources = System.Windows.Application.LoadComponent(
                    new Uri("SupportTools_Excel;component/Resources/Xaml/Application.xaml", UriKind.Relative)) as System.Windows.ResourceDictionary;

                //var resources = System.Windows.Application.LoadComponent(
                //    new Uri("pack:/SupportTools_Excel;:,,/Resources/Xaml/Application.xaml")) as System.Windows.ResourceDictionary;

                // Merge it on application level

                Common.XamlApplication.Resources.MergedDictionaries.Add(resources);

                // Wire in Prism Bootstrapper

                //try
                //{
                //    var bootstrapper = new Bootstrapper();
                //    bootstrapper.Run();
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.ToString());
                //}
            }
            catch (Exception ex)
            {
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString());
                Common.DeveloperMode = false;
            }

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void UnLoadXamlApplication()
        {
            long startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);

            try
            {
                if (null != Common.XamlApplication)
                {
                    Common.XamlApplication.Shutdown();
                    Common.XamlApplication = null;
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
    }
}
