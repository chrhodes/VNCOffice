using SupportTools_Excel.Domain;
using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

using XlHlp = VNC.AddinHelper.Excel;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;
using VNC;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    public class RequestHandlers
    {

        #region Delegates

        public delegate void ProcessAddHeaderCommand(XlHlp.XlLocation insertAt);
        public delegate void ProcessCreateWorksheetCommand(Options_AZDO_TFS options);
        public delegate void ProcessCreateWorksheetCommandSections(string sectionOptions, Options_AZDO_TFS options);
        public delegate void ProcessCreateWorksheetCommandTeamProjectCollection(string teamProjectCollectionUri, Options_AZDO_TFS options);

        #endregion

        public static void ProcessCreateWorkSheet(ProcessCreateWorksheetCommand command,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION($"Enter ({command.Method.Name})", Common.LOG_CATEGORY);

            XlHlp.DisplayInWatchWindow(string.Format("{0} {1}",
                MethodBase.GetCurrentMethod().Name,
                command.Method.Name));

            try
            {
                SpeedUpStart();
                Common.PriorStatusBar = Globals.ThisAddIn.Application.StatusBar.ToString();

                command(options);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                SpeedUpEnd();
                Globals.ThisAddIn.Application.StatusBar = Common.PriorStatusBar;
            }

            Log.APPLICATION($"Exit ({command.Method.Name})", Common.LOG_CATEGORY, startTicks);
        }

        public static void ProcessCreateWorkSheetSections(ProcessCreateWorksheetCommandSections command,
            string sectionsToDisplay, 
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION($"Enter ({command.Method.Name})", Common.LOG_CATEGORY);

            try
            {
                SpeedUpStart();
                Common.PriorStatusBar = Globals.ThisAddIn.Application.StatusBar.ToString();

                command(sectionsToDisplay, options);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                SpeedUpEnd();
                Globals.ThisAddIn.Application.StatusBar = Common.PriorStatusBar;
            }

            Log.APPLICATION($"Exit ({command.Method.Name})", Common.LOG_CATEGORY, startTicks);
        }

        public static void ProcessCreateWorkSheetTeamProjectCollection(ProcessCreateWorksheetCommandSections command,
            string teamProjectCollectionUri,
            Options_AZDO_TFS options)
        {
            long startTicks = Log.APPLICATION($"Enter ({command.Method.Name})", Common.LOG_CATEGORY);

            try
            {
                SpeedUpStart();
                Common.PriorStatusBar = Globals.ThisAddIn.Application.StatusBar.ToString();

                command(teamProjectCollectionUri, options);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                SpeedUpEnd();
                Globals.ThisAddIn.Application.StatusBar = Common.PriorStatusBar;
            }

            Log.APPLICATION($"Exit ({command.Method.Name})", Common.LOG_CATEGORY, startTicks);
        }

        public static void SpeedUpStart()
        {
            // NOTE(crhodes)
            // These keep track of prior state
            XlHlp.ScreenUpdatesOff();
            XlHlp.CalculationsOff();
        }

        public static void SpeedUpEnd()
        {
            // NOTE(crhodes)
            // These keep track of prior state, but not having screen updates make no sense
            XlHlp.CalculationsOn(force: true);
            XlHlp.ScreenUpdatesOn(force: true);
        }
    }
}
