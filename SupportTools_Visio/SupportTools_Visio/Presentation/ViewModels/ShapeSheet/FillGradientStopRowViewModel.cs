using System;

using Prism.Commands;

using SupportTools_Visio.Actions;
using SupportTools_Visio.Presentation.ModelWrappers;
using VNC;
using VNC.Core.Mvvm;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class FillGradientStopRowViewModel : ShapeSheetSectionBase
    {
        public FillGradientStopRowViewModel()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //FillGradientStopRowWrapperViewModel = new FillGradientStopRowWrapperWrapper(new Domain.FillGradientStopRowWrapperViewModel());
            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public FillGradientStopRowWrapper FillGradientStopRow { get; set; }

        public override void OnUpdateSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("UpdateFillGradientStop");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            foreach (Visio.Shape shape in selection)
            {
                //Visio_Shape.Set_FillGradientStopRowWrapperViewModel_Section(shape, FillGradientStopRowWrapperViewModel.Model);
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }

        public override void OnLoadCurrentSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Selection selection = app.ActiveWindow.Selection;

            // Verify only one shape, for now just grab first.

            foreach (Visio.Shape shape in selection)
            {
                FillGradientStopRow = new FillGradientStopRowWrapper(Visio_Shape.Get_FillGradientStopRow(shape));
                OnPropertyChanged("FillGradientStopRowWrapper");
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }
    }
}
