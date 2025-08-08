using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

using VNC;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Presentation.Views
{
    public partial class EditControlPoints : UserControl
    {
        #region Constructors and Load

        public EditControlPoints()
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            InitializeComponent();
            LoadControlContents();
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            //VNC.Log.Trace("", Common.LOG_APPNAME, 0);
            //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}()",
            //    System.Reflection.MethodInfo.GetCurrentMethod().Name));
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        private void UserControl_Unloaded(object sender, RoutedEventArgs e)
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            //VNC.Log.Trace("", Common.LOG_APPNAME, 0);
            //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}()",
            //    System.Reflection.MethodInfo.GetCurrentMethod().Name));
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        #endregion

        #region Event Handlers

        private void btnAddConnectionPoints_Click(object sender, RoutedEventArgs e)
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            //VNC.Log.Trace("", Common.LOG_APPNAME, 0);
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Common.VisioApplication.BeginUndoScope("AddConnectionPoints");

            VNCVisioToolsApplication.Actions.Visio_Shape.Add_ConnectionPoints(GetConnectionPointSettings());

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        private void btnClearConnectionPoints_Click(object sender, RoutedEventArgs e)
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            string tag = ((Button)sender).Tag.ToString();

            VNCVisioToolsApplication.Actions.Visio_Shape.ClearConnectionPoints(tag);
        }

        #endregion

        #region Private Methods

        private void LoadControlContents()
        {
            try
            {
                //visioCommand_Picker.PopulateControlFromFile(Common.cCONFIG_FILE);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #endregion

        List<VNCVisioAddIn.Domain.ConnectionPointRow> GetConnectionPointSettings()
        {
            List<VNCVisioAddIn.Domain.ConnectionPointRow> connectionPoints = new List<VNCVisioAddIn.Domain.ConnectionPointRow>();

            #region Top Edge

            if ((bool)ceTEL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "TEL",
                    X = "Width*0.0",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top,Left"
                });
            }

            if ((bool)ceT16LLL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T16LLL",
                    X = "Width*0.0625",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT8LL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T8LL",
                    X = "Width*0.125",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT16LLR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T16LLR",
                    X = "Width*0.1875",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceTQL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "TQL",
                    X = "Width*0.25",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT16LRL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T16LRL",
                    X = "Width*0.3125",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT8LR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T8LR",
                    X = "Width*0.375",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT16LRR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T16LRR",
                    X = "Width*0.4375",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceTM.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "TM",
                    X = "Width*0.5",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT16RLL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T16RLL",
                    X = "Width*0.5625",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT8RL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T8RL",
                    X = "Width*0.625",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT16RLR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T16RL",
                    X = "Width*0.6875",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceTQR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "TQR",
                    X = "Width*0.75",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT16RRL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T16RRL",
                    X = "Width*0.8125",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT8RR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T8RR",
                    X = "Width*0.875",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceT16RRR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "T16RRR",
                    X = "Width*0.9375",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top"
                });
            }

            if ((bool)ceTER.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "TER",
                    X = "Width*1.0",
                    Y = "Height*1.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Top,Right"
                });
            }

            #endregion Top

            #region Bottom

            if ((bool)ceBEL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "BEL",
                    X = "Width*0.0",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom,Left"
                });
            }

            if ((bool)ceB16LLL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B16LLL",
                    X = "Width*0.0625",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB8LL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B8LL",
                    X = "Width*0.125",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB16LLR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B16LLR",
                    X = "Width*0.1875",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceBQL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "BQL",
                    X = "Width*0.25",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB16LRL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B16LRL",
                    X = "Width*0.3125",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB8LR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B8LR",
                    X = "Width*0.375",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB16LRR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B16LRR",
                    X = "Width*0.4375",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceBM.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "BM",
                    X = "Width*0.5",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB16RLL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B16RLL",
                    X = "Width*0.5625",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB8RL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B8RL",
                    X = "Width*0.625",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB16RLR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B16RL",
                    X = "Width*0.6875",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceBQR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "BQR",
                    X = "Width*0.75",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB16RRL.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B16RRL",
                    X = "Width*0.8125",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB8RR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B8RR",
                    X = "Width*0.875",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceB16RRR.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "B16RRR",
                    X = "Width*0.9375",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom"
                });
            }

            if ((bool)ceBER.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "BER",
                    X = "Width*1.0",
                    Y = "Height*0.0",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Bottom,Right"
                });
            }

            #endregion Bottom

            #region Left

            if ((bool)ceL16TTT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L16TTT",
                    X = "Width*0.0",
                    Y = "Height*0.9375",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL8TT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L8TT",
                    X = "Width*0.0",
                    Y = "Height*0.875",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL16TTB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L16TTB",
                    X = "Width*0.0",
                    Y = "Height*0.8125",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceLQT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "LQT",
                    X = "Width*0.0",
                    Y = "Height*0.75",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL16TBT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L16TBT",
                    X = "Width*0.0",
                    Y = "Height*0.6875",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL8TB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L8TB",
                    X = "Width*0.0",
                    Y = "Height*0.625",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL16TBB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L16TBB",
                    X = "Width*0.0",
                    Y = "Height*0.5625",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceLM.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "LM",
                    X = "Width*0.0",
                    Y = "Height*0.5",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL16BTT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L16BTT",
                    X = "Width*0.0",
                    Y = "Height*0.4375",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL8BT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L8BT",
                    X = "Width*0.0",
                    Y = "Height*0.375",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL16BTB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L16BTB",
                    X = "Width*0.0",
                    Y = "Height*0.3125",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceLQB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "LQB",
                    X = "Width*0.0",
                    Y = "Height*0.25",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL16BBT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L16BBT",
                    X = "Width*0.0",
                    Y = "Height*0.1875",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL8BB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L8BB",
                    X = "Width*0.0",
                    Y = "Height*0.125",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            if ((bool)ceL16BBB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "L16BBB",
                    X = "Width*0.0",
                    Y = "Height*0.0625",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Left"
                });
            }

            #endregion Left

            #region Right

            if ((bool)ceR16TTT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R16TTT",
                    X = "Width*1.0",
                    Y = "Height*0.9375",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR8TT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R8TT",
                    X = "Width*1.0",
                    Y = "Height*0.875",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR16TTB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R16TTB",
                    X = "Width*1.0",
                    Y = "Height*0.8125",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceRQT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "RQT",
                    X = "Width*1.0",
                    Y = "Height*0.75",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR16TBT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R16TBT",
                    X = "Width*1.0",
                    Y = "Height*0.6875",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR8TB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R8TB",
                    X = "Width*1.0",
                    Y = "Height*0.625",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR16TBB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R16TBB",
                    X = "Width*1.0",
                    Y = "Height*0.5625",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceRM.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "RM",
                    X = "Width*1.0",
                    Y = "Height*0.5",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR16BTT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R16BTT",
                    X = "Width*1.0",
                    Y = "Height*0.4375",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR8BT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R8BT",
                    X = "Width*1.0",
                    Y = "Height*0.375",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR16BTB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R16BTB",
                    X = "Width*1.0",
                    Y = "Height*0.3125",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceRQB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "RQB",
                    X = "Width*1.0",
                    Y = "Height*0.25",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR16BBT.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R16BBT",
                    X = "Width*1.0",
                    Y = "Height*0.1875",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR8BB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R8BB",
                    X = "Width*1.0",
                    Y = "Height*0.125",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            if ((bool)ceR16BBB.IsChecked)
            {
                connectionPoints.Add(new VNCVisioAddIn.Domain.ConnectionPointRow
                {
                    Name = "R16BBB",
                    X = "Width*1.0",
                    Y = "Height*0.0625",
                    DirX = "0 in",
                    DirY = "0 in",
                    Type = "0",
                    D = "Right"
                });
            }

            #endregion Right

            return connectionPoints;
        }

        private void btnInitializeConnectionPoints_Click(object sender, RoutedEventArgs e)
        {
            switch (((Button)sender).Tag.ToString())
            {
                case "Tops":
                    CheckTops(true);
                    break;

                case "Bottoms":
                    CheckBottoms(true);
                    break;

                case "Lefts":
                    CheckLefts(true);
                    break;

                case "Rights":
                    CheckRights(true);
                    break;

                case "Edges":
                    CheckEdges(true);
                    break;

                case "Middles":
                    CheckMiddles(true);
                    break;

                case "Quarters":
                    CheckQuarters(true);

                    break;

                case "Eighths":
                    CheckEighths(true);
                    break;

                case "Sixteenths":
                    CheckSixteenths(true);
                    break;

                case "All":
                    CheckAll();
                    break;

                case "Clear":
                    ClearAll();
                    break;
            }
        }

        //private void btnInitializeConnectionPoints_Click(string tag)
        //{
        //    switch (tag)
        //    {
        //        case "Tops":
        //            CheckTops(false);
        //            break;

        //        case "Bottoms":
        //            CheckBottoms(false);
        //            break;

        //        case "Lefts":
        //            CheckLefts(false);
        //            break;

        //        case "Rights":
        //            CheckRights(false);
        //            break;

        //        case "All":
        //            ClearAll();
        //            break;

        //        default:
        //            MessageBox.Show($"Unknown tag: {tag}");
        //            break;
        //    }
        //}

        void ClearAll()
        {
            CheckEdges(false);
            CheckMiddles(false);
            CheckQuarters(false);
            CheckEighths(false);
            CheckSixteenths(false);
        }

        void CheckAll()
        {
            CheckEdges(true);
            CheckMiddles(true);
            CheckQuarters(true);
            CheckEighths(true);
            CheckSixteenths(true);
        }

        void CheckTops(bool isChecked)
        {
            ceTEL.IsChecked = isChecked;

            ceT16LLL.IsChecked = isChecked;
            ceT8LL.IsChecked = isChecked;
            ceT16LLR.IsChecked = isChecked;

            ceTQL.IsChecked = isChecked;

            ceT16LRL.IsChecked = isChecked;
            ceT8LR.IsChecked = isChecked;
            ceT16LRR.IsChecked = isChecked;

            ceTM.IsChecked = isChecked;

            ceT16RLL.IsChecked = isChecked;
            ceT8RL.IsChecked = isChecked;
            ceT16RLR.IsChecked = isChecked;

            ceTQR.IsChecked = isChecked;

            ceT16RRL.IsChecked = isChecked;
            ceT8RR.IsChecked = isChecked;
            ceT16RRR.IsChecked = isChecked;

            ceTER.IsChecked = isChecked;
        }

        void CheckBottoms(bool isChecked)
        {
            ceBEL.IsChecked = isChecked;

            ceB16LLL.IsChecked = isChecked;
            ceB8LL.IsChecked = isChecked;
            ceB16LLR.IsChecked = isChecked;

            ceBQL.IsChecked = isChecked;

            ceB16LRL.IsChecked = isChecked;
            ceB8LR.IsChecked = isChecked;
            ceB16LRR.IsChecked = isChecked;

            ceBM.IsChecked = isChecked;

            ceB16RLL.IsChecked = isChecked;
            ceB8RL.IsChecked = isChecked;
            ceB16RLR.IsChecked = isChecked;

            ceBQR.IsChecked = isChecked;

            ceB16RRL.IsChecked = isChecked;
            ceB8RR.IsChecked = isChecked;
            ceB16RRR.IsChecked = isChecked;

            ceBER.IsChecked = isChecked;
        }

        void CheckLefts(bool isChecked)
        {
            ceTEL.IsChecked = isChecked;

            ceL16TTT.IsChecked = isChecked;
            ceL8TT.IsChecked = isChecked;
            ceL16TTB.IsChecked = isChecked;

            ceLQT.IsChecked = isChecked;

            ceL16TBT.IsChecked = isChecked;
            ceL8TB.IsChecked = isChecked;
            ceL16TBB.IsChecked = isChecked;

            ceLM.IsChecked = isChecked;

            ceL16BTT.IsChecked = isChecked;
            ceL8BT.IsChecked = isChecked;
            ceL16BTB.IsChecked = isChecked;

            ceLQB.IsChecked = isChecked;

            ceL16BBT.IsChecked = isChecked;
            ceL8BB.IsChecked = isChecked;
            ceL16BBB.IsChecked = isChecked;

            ceBEL.IsChecked = isChecked;
        }

        void CheckRights(bool isChecked)
        {
            ceTER.IsChecked = isChecked;

            ceR16TTT.IsChecked = isChecked;
            ceR8TT.IsChecked = isChecked;
            ceR16TTB.IsChecked = isChecked;

            ceRQT.IsChecked = isChecked;

            ceR16TBT.IsChecked = isChecked;
            ceR8TB.IsChecked = isChecked;
            ceR16TBB.IsChecked = isChecked;

            ceRM.IsChecked = isChecked;

            ceR16BTT.IsChecked = isChecked;
            ceR8BT.IsChecked = isChecked;
            ceR16BTB.IsChecked = isChecked;

            ceRQB.IsChecked = isChecked;

            ceR16BBT.IsChecked = isChecked;
            ceR8BB.IsChecked = isChecked;
            ceR16BBB.IsChecked = isChecked;

            ceBER.IsChecked = isChecked;
        }

        private void CheckSixteenths(bool isChecked)
        {
            ceT16LLL.IsChecked = isChecked;
            ceT16LLR.IsChecked = isChecked;
            ceT16LRL.IsChecked = isChecked;
            ceT16LRR.IsChecked = isChecked;
            ceT16RLL.IsChecked = isChecked;
            ceT16RLR.IsChecked = isChecked;
            ceT16RRL.IsChecked = isChecked;
            ceT16RRR.IsChecked = isChecked;

            ceB16LLL.IsChecked = isChecked;
            ceB16LLR.IsChecked = isChecked;
            ceB16LRL.IsChecked = isChecked;
            ceB16LRR.IsChecked = isChecked;
            ceB16RLL.IsChecked = isChecked;
            ceB16RLR.IsChecked = isChecked;
            ceB16RRL.IsChecked = isChecked;
            ceB16RRR.IsChecked = isChecked;

            ceL16TTT.IsChecked = isChecked;
            ceL16TTB.IsChecked = isChecked;
            ceL16TBT.IsChecked = isChecked;
            ceL16TBB.IsChecked = isChecked;
            ceL16BTT.IsChecked = isChecked;
            ceL16BTB.IsChecked = isChecked;
            ceL16BBT.IsChecked = isChecked;
            ceL16BBB.IsChecked = isChecked;

            ceR16TTT.IsChecked = isChecked;
            ceR16TTB.IsChecked = isChecked;
            ceR16TBT.IsChecked = isChecked;
            ceR16TBB.IsChecked = isChecked;
            ceR16BTT.IsChecked = isChecked;
            ceR16BTB.IsChecked = isChecked;
            ceR16BBT.IsChecked = isChecked;
            ceR16BBB.IsChecked = isChecked;
        }

        private void CheckEighths(bool isChecked)
        {
            ceT8LL.IsChecked = isChecked;
            ceT8LR.IsChecked = isChecked;
            ceT8RL.IsChecked = isChecked;
            ceT8RR.IsChecked = isChecked;

            ceB8LL.IsChecked = isChecked;
            ceB8LR.IsChecked = isChecked;
            ceB8RL.IsChecked = isChecked;
            ceB8RR.IsChecked = isChecked;

            ceL8TT.IsChecked = isChecked;
            ceL8TB.IsChecked = isChecked;
            ceL8BT.IsChecked = isChecked;
            ceL8BB.IsChecked = isChecked;

            ceR8TT.IsChecked = isChecked;
            ceR8TB.IsChecked = isChecked;
            ceR8BT.IsChecked = isChecked;
            ceR8BB.IsChecked = isChecked;
        }

        private void CheckQuarters(bool isChecked)
        {
            ceTQL.IsChecked = isChecked;
            ceTQR.IsChecked = isChecked;

            ceBQL.IsChecked = isChecked;
            ceBQR.IsChecked = isChecked;

            ceLQT.IsChecked = isChecked;
            ceLQB.IsChecked = isChecked;

            ceRQT.IsChecked = isChecked;
            ceRQB.IsChecked = isChecked;
        }

        private void CheckMiddles(bool isChecked)
        {
            ceTM.IsChecked = isChecked;
            ceBM.IsChecked = isChecked;

            ceLM.IsChecked = isChecked;
            ceRM.IsChecked = isChecked;
        }

        private void CheckEdges(bool isChecked)
        {
            ceTEL.IsChecked = isChecked;
            ceTER.IsChecked = isChecked;

            ceBEL.IsChecked = isChecked;
            ceBER.IsChecked = isChecked;
        }
    }
}
