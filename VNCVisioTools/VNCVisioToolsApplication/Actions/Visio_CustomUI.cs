using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;

using VNCVisioToolsApplication.Presentation.ViewModels;
using VNCVisioToolsApplication.Presentation.Views;

namespace VNCVisioToolsApplication.Actions
{
    public class Visio_CustomUI
    {

        #region CommandCockpit

        public static DxThemedWindowHost commandCockpitHost = null;

        public static void CommandCockpit()
        {
            if (commandCockpitHost is null) commandCockpitHost = new DxThemedWindowHost(Common.EventAggregator);

            commandCockpitHost.DisplayUserControlInHost(
                "Command Cockpit (XML Commands)",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                //(CommandCockpitViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpitViewModel))
                (CommandCockpit)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpit))
            );
        }

        //public static DxThemedWindowHost linqToExcelHost = null;

        //public static void Linq2Excel()
        //{
        //    if (linqToExcelHost is null) linqToExcelHost = new DxThemedWindowHost();

        //    linqToExcelHost.DisplayUserControlInHost(
        //        "Linq to Excel",
        //        Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
        //        ShowWindowMode.Modeless_Show,
        //        (Presentation.Views.LinqToExcel)Common.ApplicationBootstrapper.Container.Resolve(typeof(Presentation.Views.LinqToExcel))
        //    );
        //}

        #endregion

        #region DuplictaePage

        public static DxThemedWindowHost duplicatePageHost = null;

        public static void DuplicatePage()
        {
            if (duplicatePageHost is null) duplicatePageHost = new DxThemedWindowHost(Common.EventAggregator);

            duplicatePageHost.DisplayUserControlInHost(
            "Duplicate Page",
            Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
            //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            ShowWindowMode.Modeless_Show,
            new Presentation.Views.DuplicatePage(new DuplicatePageViewModel()));
        }

        #endregion

        #region RenamePages

        public static DxThemedWindowHost renamePagesHost = null;

        public static void RenamePages()
        {
            if (renamePagesHost is null) renamePagesHost = new DxThemedWindowHost(Common.EventAggregator);

            renamePagesHost.DisplayUserControlInHost(
                "Rename Paqe(s)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (RenamePageViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(RenamePageViewModel))
            );
        }

        #endregion

        #region MovePages

        public static DxThemedWindowHost movePagesHost = null;

        public static void MovePages()
        {
            if (movePagesHost is null) movePagesHost = new DxThemedWindowHost(Common.EventAggregator);

            movePagesHost.DisplayUserControlInHost(
                "Move Paqe(s)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (MovePageViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(MovePageViewModel))
            );
        }

        #endregion

        #region EditControlRows

        public static DxThemedWindowHost editControlRowsHost = null;

        public static void EditControlRows()
        {
            if (editControlRowsHost is null) editControlRowsHost = new DxThemedWindowHost(Common.EventAggregator);

            editControlRowsHost.DisplayUserControlInHost(
                "Edit Control Rows",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.EditControlRows(new EditControlRowsViewModel()));
        }

        #endregion

        #region EditParagraph

        public static DxThemedWindowHost editParagraphHost = null;

        public static void EditParagraph()
        {
            if (editParagraphHost is null) editParagraphHost = new DxThemedWindowHost(Common.EventAggregator);

            editParagraphHost.DisplayUserControlInHost(
                "Edit Paragraph",
                300, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.EditParagraph(new EditParagraphViewModel()));
        }

        #endregion

        #region EditControlPoints

        public static DxThemedWindowHost editControlPointsHost = null;

        public static void EditControlPoints()
        {
            if (editControlPointsHost is null) editControlPointsHost = new DxThemedWindowHost(Common.EventAggregator);

            editControlPointsHost.DisplayUserControlInHost(
                "Edit Shape Control Points Text",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new EditControlPoints());
        }

        #endregion

        #region EditText

        public static DxThemedWindowHost editTextHost = null;

        public static void EditText()
        {
            if (editTextHost is null) editTextHost = new DxThemedWindowHost(Common.EventAggregator);

            editTextHost.DisplayUserControlInHost(
                "Edit Text",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new EditText(new EditTextViewModel()));
        }

        #endregion

        #region CustomUI_Car

        public static DxThemedWindowHost carMainHost = null;

        public static void CustomUI_Car()
        {
            Common.WriteToDebugWindow("CustomUI_Car() - Not Implemented", true);
            // TODO(crhodes)
            // Decide if want to keep this

            //if (carMainHost is null) carMainHost = new DxThemedWindowHost();

            //carMainHost.DisplayUserControlInHost(
            //    "CarMain",
            //    Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            //    ShowWindowMode.Modeless_Show,
            //    (Explore.Presentation.Views.CarMain)Common.ApplicationBootstrapper.Container.Resolve(typeof(Explore.Presentation.Views.CarMain))
            //);
        }

        #endregion

        #region TestVisioLogging

        public static DxThemedWindowHost TestVisioLoggingHost = null;

        public static void TestVisioLogging()
        {
            if (TestVisioLoggingHost is null) TestVisioLoggingHost = new DxThemedWindowHost(Common.EventAggregator);

            TestVisioLoggingHost.DisplayUserControlInHost(
                "Folder Map",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                //(CommandCockpitViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpitViewModel))
                (TestVisioLogging)Common.ApplicationBootstrapper.Container.Resolve(typeof(TestVisioLogging))
            //new TestVisioLogging()
            );
        }

        #endregion

        #region LoggingConfiguration

        public static DxThemedWindowHost LoggingConfigurationHost = null;

        public static void LoggingConfiguration()
        {
            if (LoggingConfigurationHost is null) LoggingConfigurationHost = new DxThemedWindowHost(Common.EventAggregator);

            LoggingConfigurationHost.DisplayUserControlInHost(
                "Folder Map",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                //(CommandCockpitViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpitViewModel))
                //(LoggingConfiguration)Common.ApplicationBootstrapper.Container.Resolve(typeof(LoggingConfiguration))
                new VNCLoggingConfigMain()
            );
        }

        #endregion
    }
}
