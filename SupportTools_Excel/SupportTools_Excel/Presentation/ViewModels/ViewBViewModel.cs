﻿using System;

using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Excel.Presentation.ViewModels
{
    public class ViewBViewModel : EventViewModelBase, IViewBViewModel, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public ViewBViewModel(
            IEventAggregator eventAggregator,
            DialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            //

            Message = "ViewBViewModel says hello";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

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

        private string _messageA;

        public string MessageA
        {
            get => _messageA;
            set
            {
                if (_messageA == value)
                    return;
                _messageA = value;
                OnPropertyChanged();
            }
        }

        private string _messageB;

        public string MessageB
        {
            get => _messageB;
            set
            {
                if (_messageB == value)
                    return;
                _messageB = value;
                OnPropertyChanged();
            }
        }

        private string _messageC;

        public string MessageC
        {
            get => _messageC;
            set
            {
                if (_messageC == value)
                    return;
                _messageC = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Event Handlers


        #endregion

        #region Public Methods


        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods


        #endregion

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
