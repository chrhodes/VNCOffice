using System;

using SupportTools_Visio.Actions;
using SupportTools_Visio.Presentation.ModelWrappers;

using VNC;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class FillFormatViewModel : ShapeSheetSectionBase
    {
        public FillFormatViewModel() : base()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            UpdateButtonContent = "Update FillFormat for selected shapes";
            // TODO(crhodes)
            // Decide if we want defaults
            //FillFormatViewModel = new FillFormatWrapper(new Domain.FillFormatViewModel());

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public FillFormatWrapper FillFormat { get; set; }

        public override void OnUpdateSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("UpdateFillFormat");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Visio_Shape.Set_FillFormat_Section(shape, FillFormat.Model);
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }

        public override void OnLoadCurrentSettingsExecute()
        {

            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                FillFormat = new FillFormatWrapper(Visio_Shape.Get_FillFormat(shape));
                OnPropertyChanged("FillFormat");
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }
    }
}
