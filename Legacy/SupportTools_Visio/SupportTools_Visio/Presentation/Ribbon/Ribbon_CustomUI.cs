using System;

using Microsoft.Office.Tools.Ribbon;

using SupportTools_Visio.Presentation.ViewModels;
using SupportTools_Visio.Presentation.Views;

using VNC;
using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;

namespace SupportTools_Visio
{
    public partial class Ribbon
    {
        #region Event Handlers

        #region WPF Events - Custom

        public static DxThemedWindowHost commandCockpitHost = null;

        private void btnCommandCockpit_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (commandCockpitHost is null) commandCockpitHost = new DxThemedWindowHost();

            commandCockpitHost.DisplayUserControlInHost(
                "Command Cockpit (XML Commands)",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                //(CommandCockpitViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpitViewModel))
                (CommandCockpit)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpit))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost linqToExcelHost = null;

        private void btnLinqToExcel_Click
            (object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (linqToExcelHost is null) linqToExcelHost = new DxThemedWindowHost();

            linqToExcelHost.DisplayUserControlInHost(
                "Linq to Excel",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                (Presentation.Views.LinqToExcel)Common.ApplicationBootstrapper.Container.Resolve(typeof(Presentation.Views.LinqToExcel))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost duplicatePageHost = null;

        private void btnDuplicatePage_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (duplicatePageHost is null) duplicatePageHost = new DxThemedWindowHost();

            duplicatePageHost.DisplayUserControlInHost(
            "Duplicate Page",
            Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
            //Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            ShowWindowMode.Modeless_Show,
            new Presentation.Views.DuplicatePage(new DuplicatePageViewModel()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost renamePagesHost = null;

        private void btnRenamePages_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (renamePagesHost is null) renamePagesHost = new DxThemedWindowHost();

            renamePagesHost.DisplayUserControlInHost(
                "Rename Paqe(s)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (RenamePageViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(RenamePageViewModel))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost movePagesHost = null;

        private void btnMovePages_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (movePagesHost is null) movePagesHost = new DxThemedWindowHost();

            movePagesHost.DisplayUserControlInHost(
                "Move Paqe(s)",
                Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
                ShowWindowMode.Modeless_Show,
                (MovePageViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(MovePageViewModel))
            );

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost editControlRowsHost = null;

        private void btnEditControlRows_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (editControlRowsHost is null) editControlRowsHost = new DxThemedWindowHost();

            editControlRowsHost.DisplayUserControlInHost(
                "Edit Control Rows",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.EditControlRows(new EditControlRowsViewModel()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost editParagraphHost = null;

        private void btnEditParagraph_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (editParagraphHost is null) editParagraphHost = new DxThemedWindowHost();

            editParagraphHost.DisplayUserControlInHost(
                "Edit Paragraph",
                300, 600,
                ShowWindowMode.Modeless_Show,
                new Presentation.Views.EditParagraph(new EditParagraphViewModel()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //private EditControlPoints editControlPointsUC = null;

        //private Presentation.Views.EditText editTextUC = null;

        public static DxThemedWindowHost editControlPointsHost = null;

        private void btnEditControlPoints_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (editControlPointsHost is null) editControlPointsHost = new DxThemedWindowHost();

            editControlPointsHost.DisplayUserControlInHost(
                "Edit Shape Control Points Text",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new EditControlPoints());

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost editTextHost = null;

        private void btnEditText_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            if (editTextHost is null) editTextHost = new DxThemedWindowHost();

            editTextHost.DisplayUserControlInHost(
                "Edit Text",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                new EditText(new EditTextViewModel()));

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static DxThemedWindowHost carMainHost = null;

        private void btnCustomUI_Car_Click(object sender, RibbonControlEventArgs e)
        {
            //DxThemedWindowHost.DisplayUserControlInHost(ref carMainHost,
            //    "CarMain",
            //    Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
            //    ShowWindowMode.Modeless_Show,
            //    (Explore.Presentation.ViewModels.CarMainViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(Explore.Presentation.ViewModels.CarMainViewModel))
            //);

            if (carMainHost is null) carMainHost = new DxThemedWindowHost();

            carMainHost.DisplayUserControlInHost(
                "CarMain",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                (Explore.Presentation.Views.CarMain)Common.ApplicationBootstrapper.Container.Resolve(typeof(Explore.Presentation.Views.CarMain))
            );

            //DxThemedWindowHost.DisplayUserControlInHost(ref carMainHost,
            //    "ViewABC",
            //    Common.DEFAULT_WINDOW_WIDTH_SMALL, Common.DEFAULT_WINDOW_HEIGHT_SMALL,
            //    ShowWindowMode.Modeless_Show,
            //    (Explore.Presentation.Views.ViewABC)Common.ApplicationBootstrapper.Container.Resolve(typeof(Explore.Presentation.Views.ViewABC))
            //);
        }

        #endregion

        #region Document Base Pages

        private void btnAddArchitectureBasePages_Click(object sender, RibbonControlEventArgs e)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Actions.Visio_Document.AddArchitectureBasePages();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #endregion Event Handlers
    }
}