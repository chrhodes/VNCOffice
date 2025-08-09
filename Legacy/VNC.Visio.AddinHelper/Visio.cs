using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.AddinHelper
{
    public class Visio
	{
		private Microsoft.Office.Interop.Visio.Application _VisioApplication;

		public Microsoft.Office.Interop.Visio.Application VisioApplication
		{
			get { return _VisioApplication; }
			set { _VisioApplication = value; }
		}

		private bool _enableScreenUpdatesToggle = true;

		public bool EnableScreenUpdatesToggle
		{
			get { return _enableScreenUpdatesToggle; }
			set { _enableScreenUpdatesToggle = value; }
		}

		public static void DisplayInWatchWindow(string outputLine)
		{
            Log.APPLICATION($"{outputLine}", Common.LOG_CATEGORY);
			Common.WriteToWatchWindow($"{outputLine}");
		}

        public static bool LoadStencil(Microsoft.Office.Interop.Visio.Application app, string stencilName)
        {
            bool result = false;

            try
            {
                var foo = app.Documents[stencilName];
                result = true;
            }
            catch (Exception)
            {
                // Stencil may not be open.  Try opening it

                try
                {
                    app.Documents.OpenEx(stencilName, (short)VisOpenSaveArgs.visOpenRO + (short)VisOpenSaveArgs.visOpenDocked);
                    result = true;
                }
                catch (Exception)
                {
                    MessageBox.Show($"Cannot locate or open {stencilName}, aborting.");
                }
            }

            return result;
        }
    }
}
