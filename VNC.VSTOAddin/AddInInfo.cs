using System;
using System.Windows.Forms;
using System.Reflection;
//using Microsoft.Office.Core;

namespace VNC.VSTOAddIn
{
    /// <summary>
    /// AddInInfo
    /// </summary>
    /// <remarks>
    /// This class can be used in two ways.  If calling this from a commandBar, modify
    /// the Private Const as needed and then create an instance of this class in the code
    /// that loads the command bars.
    /// 
    /// If calling this from a Ribbon Event handler, call the ActionNameGoesHere method directly.
    /// 
    /// Rename the ActionNameGoesHere Method and provide code that does something useful.
    /// </remarks>
    public class AddInInfo     {

        #region "Private Constants and Variables"

        private const string _MODULE_NAME = Common.LOG_CATEGORY + "AddInInfo";
        private const string _NAME = "AddInInfo";
        private const string _BITMAP_NAME = "AddinInfo.bmp";
        private const string _CAPTION = "AddInInfo";
        private const string _TOOL_TIP_TEXT = "Click for AddInInfo";
        private const string _DESCRIPTION = "AddInInfo does ...";
        #endregion

        #region "Public Methods"

        public static void DisplayInfo()
        {
            //AssemblyHelper.AssemblyInformation info = new AssemblyHelper.AssemblyInformation(System.Reflection.Assembly.GetExecutingAssembly());
            VNC.AssemblyHelper.AssemblyInformation info = new VNC.AssemblyHelper.AssemblyInformation(System.Reflection.Assembly.GetCallingAssembly());
            MessageBox.Show(info.ToString());
        }

        #endregion

    }
}