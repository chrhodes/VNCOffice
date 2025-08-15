using System;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Resources;

using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;

namespace DemoAndTestLogging1.Presentation.Views
{
    public partial class StatusBar : ViewBase, IStatusBar
    {
        #region Constructors, Initialization, and Load

        public StatusBar()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;

            InitializeComponent();

            // Wire up ViewModel if needed

            // If View First with ViewModel in Xaml

            // ViewModel = (IStatusBarViewModel)DataContext;

            // Can create directly

            // ViewModel = new StatusBarViewModel();

            // Can use ourselves for everything

            //DataContext = this;

            InitializeView();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public StatusBar(ViewModels.IStatusBarViewModel viewModel)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Constructor) startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountVP++;

            InitializeComponent();

            ViewModel = viewModel;  // ViewBase sets the DataContext to ViewModel

            // For the rare case where the ViewModel needs to know about the View
            // TODO(crhodes)
            // Maybe the button image logic should be moved into View and out of ViewModel
            ViewModel.View = this;  // Need this to get back to button images

            InitializeView();

            if (Common.VNCLogging.Constructor) Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeView()
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.ViewLow) startTicks = Log.VIEW_LOW("Enter", Common.LOG_CATEGORY);

            // Store information about the View, DataContext, and ViewModel 
            // for the DeveloperInfo control. Useful for debugging binding issues
            // Set the DataConext to us.

            ViewType = this.GetType().ToString().Split('.').Last();
            ViewModelType = ViewModel?.GetType().ToString().Split('.').Last();
            ViewDataContextType = this.DataContext?.GetType().ToString().Split('.').Last();
            spDeveloperInfo.DataContext = this;

            // TODO(crhodes)
            // Put things here that initialize the View
            // Hook eventhandlers, etc.

            eventAggregator = (IEventAggregator)Common.Container.Resolve(typeof(IEventAggregator));
            statusMessageEvent = eventAggregator.GetEvent<StatusMessageEvent>();
            developerModeEvent = eventAggregator.GetEvent<DeveloperModeEvent>();

#if DEBUG
            // Xaml uses VNCDeveloperMotivation
            DeveloperModeToolTip = "Turn off Developer Mode";
#else
            Uri resourceUri = new Uri(".\\Resources\\Images\\ToolBox.png", UriKind.Relative);
            StreamResourceInfo streamInfo = App.GetResourceStream(resourceUri);
            BitmapFrame bitmapFrame = BitmapFrame.Create(streamInfo.Stream);

            ImageBrush brush = new ImageBrush();
            brush.ImageSource = bitmapFrame;
            btnDeveloperMode.Background = brush;

            btnDeveloperMode.Width = 24;
            btnDeveloperMode.Height = 20;

            DeveloperModeToolTip = "Turn on Developer Mode";
#endif

            // Establish any additional DataContext(s), e.g. to things held in this View

            btnDeveloperMode.DataContext = this;

            if (Common.VNCLogging.ViewLow) Log.VIEW_LOW("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)



        #endregion

        #region Structures (none)



        #endregion

        #region Fields and Properties

        public string DeveloperModeContent { get; set; } = "DeveloperMode";

        private string? _developerModeToolTip = "DeveloperMode ToolTip";
        public string? DeveloperModeToolTip
        {
            get => _developerModeToolTip;
            set
            {
                if (_developerModeToolTip == value) { return; }

                _developerModeToolTip = value;
                OnPropertyChanged();
            }
        }

        IEventAggregator eventAggregator;
        DeveloperModeEvent developerModeEvent;
        StatusMessageEvent statusMessageEvent;

        #endregion

        #region Event Handlers

        private void btnDeveloperMode_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.EventHandler) startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            IEventAggregator eventAggregator = (IEventAggregator)Common.Container.Resolve(typeof(IEventAggregator));
            var statusMessageEvent = eventAggregator.GetEvent<StatusMessageEvent>();

            if (Common.DeveloperMode)
            {
                PublishStatusMessage("Cool, you turned off DeveloperMode");

                DeveloperModeToolTip = "Turn on Developer Mode";

                if (Common.CurrentShell is not null) Common.CurrentShell.DeveloperUIMode = System.Windows.Visibility.Collapsed;

                // NOTE(crhodes)
                // Use this if Build Action - Content - Copy
                //ImageBrush brush = new ImageBrush();
                //brush.ImageSource = new BitmapImage(new Uri(".\\Resources\\Images\\ToolBox.png", UriKind.Relative));
                //((StatusBar)View).btnDeveloperMode.Background = brush;

                // NOTE(crhodes)
                // Use this if Build Action - Resource
                Uri resourceUri = new Uri(".\\Resources\\Images\\ToolBox.png", UriKind.Relative);
                StreamResourceInfo streamInfo = App.GetResourceStream(resourceUri);
                BitmapFrame bitmapFrame = BitmapFrame.Create(streamInfo.Stream);

                ImageBrush brush = new ImageBrush();
                brush.ImageSource = bitmapFrame;
                btnDeveloperMode.Background = brush;

                btnDeveloperMode.Width = 24;
                btnDeveloperMode.Height = 20;
            }
            else
            {
                PublishStatusMessage("Cool, you turned on DeveloperMode.  Use your power for good!");
                DeveloperModeToolTip = "Turn off Developer Mode";

                if (Common.CurrentShell is not null) Common.CurrentShell.DeveloperUIMode = System.Windows.Visibility.Visible;

                // NOTE(crhodes)
                //// Use this if Build Action - Content - Copy
                //ImageBrush brush = new ImageBrush();
                //brush.ImageSource = new BitmapImage(new Uri(".\\Resources\\Images\\VNCDeveloperMotivation.png", UriKind.Relative));
                //((StatusBar)View).btnDeveloperMode.Background = brush;

                // NOTE(crhodes)
                // Use this if Build Action - Resource
                Uri resourceUri = new Uri(".\\Resources\\Images\\VNCDeveloperMotivation.png", UriKind.Relative);
                StreamResourceInfo streamInfo = App.GetResourceStream(resourceUri);
                BitmapFrame bitmapFrame = BitmapFrame.Create(streamInfo.Stream);

                ImageBrush brush = new ImageBrush();
                brush.ImageSource = bitmapFrame;
                btnDeveloperMode.Background = brush;

                btnDeveloperMode.Width = 48;
                btnDeveloperMode.Height = 39;
            }

            Common.DeveloperMode = !Common.DeveloperMode;

            PublishDeveloperMode(Common.DeveloperMode);

            if (Common.VNCLogging.EventHandler) Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Commands (none)



        #endregion

        #region Public Methods (none)



        #endregion

        #region Protected Methods (none)



        #endregion

        #region Private Methods

        // NOTE(crhodes)
        // Publish DeveloperModeEvent for things that don't have access to Common.DeveloperMode
        // e.g. Host Windows
        private void PublishDeveloperMode(Boolean developerMode)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Event) startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            developerModeEvent.Publish(developerMode);

            if (Common.VNCLogging.Event) Log.EVENT("Enter", Common.LOG_CATEGORY, startTicks);
        }

        private void PublishStatusMessage(string message)
        {
            Int64 startTicks = 0;
            if (Common.VNCLogging.Event) startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            statusMessageEvent.Publish(message);

            if (Common.VNCLogging.Event) Log.EVENT("Enter", Common.LOG_CATEGORY, startTicks);
        }

        #endregion
    }
}
