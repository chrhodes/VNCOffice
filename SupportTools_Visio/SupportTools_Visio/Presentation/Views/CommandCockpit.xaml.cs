using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.Views
{
    public partial class CommandCockpit : ViewBase, IInstanceCountV
    {

        public CommandCockpit()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            //LoadControlContents();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public CommandCockpit(ViewModels.ICommandCockpitViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()}", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            ViewModel = viewModel;

            //LoadControlContents();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region IInstanceCount

        private static int _instanceCountV;

        public int InstanceCountV
        {
            get => _instanceCountV;
            set => _instanceCountV = value;
        }

        private static int _instanceCountVP;

        public int InstanceCountVP
        {
            get => _instanceCountVP;
            set => _instanceCountVP = value;
        }

        #endregion

        #region Private Methods

        //private void LoadControlContents()
        //{
        //    try
        //    {
        //        visioCommand_Picker.PopulateControlFromFile(Common.cCONFIG_FILE);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

        #endregion

        //public void visioCommand_Picker_ControlChanged()
        //{
        //    //var command = visioCommand_Picker.Command;
        //    //teCommandElements.Text = command.ToString();
        //    //ParseCommand(command);
        //}


    }
}
