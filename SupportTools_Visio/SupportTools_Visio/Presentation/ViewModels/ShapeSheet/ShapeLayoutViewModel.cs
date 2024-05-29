using System;

using SupportTools_Visio.Actions;
using SupportTools_Visio.Presentation.ModelWrappers;

using VNC;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class ShapeLayoutViewModel : ShapeSheetSectionBase
    {
        public ShapeLayoutWrapper ShapeLayout { get; set; }

        public ShapeLayoutViewModel() : base()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            UpdateButtonContent = "Update ShapeLayout for selected shapes";
            // TODO(crhodes)
            // Decide if we want defaults
            //ShapeLayoutViewModel = new ShapeLayoutWrapper(new Domain.ShapeLayoutViewModel());

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public override void OnUpdateSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("UpdateShapeLayout");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                Visio_Shape.Set_ShapeLayout_Section(shape, ShapeLayout.Model);
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
                ShapeLayout = new ShapeLayoutWrapper(Visio_Shape.Get_ShapeLayout(shape));
                OnPropertyChanged("ShapeLayout");
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }
    }
}
