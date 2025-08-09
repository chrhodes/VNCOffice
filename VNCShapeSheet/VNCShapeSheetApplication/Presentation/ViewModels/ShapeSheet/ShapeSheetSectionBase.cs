using System;

using Prism.Commands;

using VNCShapeSheetApplication.Core;

using VNC;
using VNC.Core.Mvvm;

using MSVisio = Microsoft.Office.Interop.Visio;

namespace VNCShapeSheetApplication.Presentation.ViewModels
{
    public class ShapeSheetSectionBase : ViewModelBase
    {
        public DelegateCommand UpdateSettings { get; protected set; }
        public DelegateCommand LoadCurrentSettings { get; protected set; }
        public DelegateCommand ExportSettings { get; protected set; }

        public DelegateCommand Refresh { get; protected set; }

        public ShapeSheetSectionBase()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            UpdateSettings = new DelegateCommand(OnUpdateSettingsExecute, OnUpdateSettingsCanExecute);
            LoadCurrentSettings = new DelegateCommand(OnLoadCurrentSettingsExecute, OnLoadCurrentSettingsCanExecute);
            ExportSettings = new DelegateCommand(ExportSettingsExecute, ExportSettingsCanExecute);

            Common.EventAggregator.GetEvent<SelectionChangedEvent>().Subscribe(OnRefresh);

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        string _message = "";
        public string Message
        {
            get { return _message; }
            set
            {
                _message = value;
                OnPropertyChanged();
            }
        }

        protected void OnRefresh()
        {
            LoadCurrentSettings.RaiseCanExecuteChanged();
        }

        string _LoadButtonContent = "Load from Current Shape";
        public string LoadButtonContent
        {
            get { return _LoadButtonContent; }
            set
            {
                _LoadButtonContent = value;
                OnPropertyChanged();
            }
        }

        string _UpdateButtonContent = "Update Items";
        public string UpdateButtonContent
        {
            get { return _UpdateButtonContent; }
            set
            {
                _UpdateButtonContent = value;
                OnPropertyChanged();
            }
        }


        string _exportSettingsContent = "Export Rows";
        public string ExportSettingsContent
        {
            get { return _exportSettingsContent; }
            set
            {
                _exportSettingsContent = value;
                OnPropertyChanged();
            }
        }

        int _selectedShapeCount;
        public int SelectedShapeCount
        {
            get { return _selectedShapeCount; }
            set
            {
                if (_selectedShapeCount == value)
                    return;
                _selectedShapeCount = value;
                OnPropertyChanged();
                //LoadCurrentSettings.RaiseCanExecuteChanged();
            }
        }

        public virtual void OnLoadCurrentSettingsExecute()
        {
            // TODO(crhodes)
            // Validate we have new settings

            Message = "OnLoadCurrentSettingsExecute Called";
        }

        public virtual Boolean OnLoadCurrentSettingsCanExecute()
        {
            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Selection selection = app.ActiveWindow.Selection;

            var count = selection.Count;
            SelectedShapeCount = count;

            if (count > 0)
            {
                if (count != 1)
                {
                    LoadButtonContent = "Must select single shape to load settings";

                    return false;
                }

                LoadButtonContent = "Load from Current Shape";

                return true;
            }
            else
            {
                LoadButtonContent = "No shape selected";

                return false;
            }
        }

        public virtual void OnUpdateSettingsExecute()
        {
            Message = "OnLoadCurrentSettingsExecute Called";
        }

        public virtual Boolean OnUpdateSettingsCanExecute()
        {
            return true;
        }

        public virtual void ExportSettingsExecute()
        {
            Message = "ExportSettingsExecute Called";
        }

        public virtual Boolean ExportSettingsCanExecute()
        {
            return true;
        }
    }
}
