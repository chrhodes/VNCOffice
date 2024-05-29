using System;

using Prism.Commands;

using SupportTools_Visio.Actions;
using SupportTools_Visio.Presentation.ModelWrappers;
using VNC;
using VNC.Core.Mvvm;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class TextFieldRowViewModel : ShapeSheetSectionBase
    {
        public TextFieldRowViewModel()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Decide if we want defaults
            //TextFieldRowViewModel = new TextFieldRowWrapper(new Domain.TextFieldRowViewModel());

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public TextFieldRowWrapper TextFieldRow { get; set; }

        public override void OnUpdateSettingsExecute()
        {
            Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!
            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("UpdateControlRow");

            // Just need to pass in the model.

            //Visio.Application app = Globals.ThisAddIn.Application;

            //Visio.Selection selection = app.ActiveWindow.Selection;

            //// Verify only one shape, for now just grab first.

            //foreach (Visio.Shape shape in selection)
            //{
            //    //Visio_Shape.Set_TextFieldRowViewModel_Section(shape, TextFieldRowViewModel.Model);
            //}

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
                TextFieldRow = new TextFieldRowWrapper(Visio_Shape.Get_TextFieldRow(shape));
                OnPropertyChanged("TextFieldRow");
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY);
        }
    }
}
