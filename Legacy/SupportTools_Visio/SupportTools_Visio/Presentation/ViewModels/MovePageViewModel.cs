using System;
using System.Collections.ObjectModel;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using SupportTools_Visio.Presentation.Views;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class MovePageViewModel : EventViewModelBase, IMovePageViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public MovePageViewModel(
            IEventAggregator eventAggregator,
            DialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public MovePageViewModel(
            MovePage view,
            IEventAggregator eventAggregator,
            DialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            View = view;
            View.ViewModel = this;
            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            MovePagesCommand = new DelegateCommand(MovePages, MovePagesCanExecute);

            var openDocs = Actions.Visio_Application.GetOpenDocuments();

            ObservableCollection<string> drawings = new ObservableCollection<string>();

            var activeDocumentName = Actions.Visio_Application.GetActiveDocument().Name;

            foreach (var doc in openDocs)
            {
                if (! doc.Name.Contains(".vssx"))
                {
                    // TODO(crhodes)
                    // Filter out Active Document
                    if (doc.Name != activeDocumentName)
                    {
                        drawings.Add(doc.Name);
                    }
                }
            }

            OpenDrawings = drawings;

            Message = "MovePageViewModel says hello";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Constructors, Initialization, and Load

        #region Fields and Properties


        private string _message;

        public string Message
        {
            get => _message;
            set
            {
                if (_message == value)
                    return;
                _message = value;
                OnPropertyChanged();
            }
        }

        private string _selectedDrawing;
        public string SelectedDrawing
        {
            get => _selectedDrawing;
            set
            {
                if (_selectedDrawing == value)
                    return;
                _selectedDrawing = value;
                OnPropertyChanged();
            }
        }

        private ObservableCollection<string> _openDrawings;

        public ObservableCollection<string> OpenDrawings
        {
            get => _openDrawings;
            set
            {
                if (_openDrawings == value)
                    return;
                _openDrawings = value;
                OnPropertyChanged();
            }
        }
        

        #endregion Fields and Properties

        #region Event Handlers

        #region MovePages Command

        public DelegateCommand MovePagesCommand { get; set; }
        public string MovePagesContent { get; set; } = "MovePages";
        public string MovePagesToolTip { get; set; } = "MovePages ToolTip";

        // Can get fancy and use Resources
        //public string MovePagesContent { get; set; } = "ViewName_MovePagesContent";
        //public string MovePagesToolTip { get; set; } = "ViewName_MovePagesContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_MovePagesContent">MovePages</system:String>
        //    <system:String x:Key="ViewName_MovePagesContentToolTip">MovePages ToolTip</system:String>

        public void MovePages()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called MovePages";

            Actions.Visio_Document.MovePages(SelectedDrawing);

            // Uncomment this if you are telling someone else to handle this
            // Common.EventAggregator.GetEvent<MovePagesEvent>().Publish();

            // Start Cut Three - Put this in PrismEvents

            //public class MovePagesEvent : PubSubEvent { }

            // End Cut Three

            // Start Cut Four - Put this in places that listen for event

            //Common.EventAggregator.GetEvent<MovePagesEvent>().Subscribe(MovePages);

            // End Cut Four
            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool MovePagesCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion MovePages Command

        // End Cut One

        #endregion Event Handlers

        #region Private Methods

        #endregion Private Methods

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion IInstanceCount
    }
}