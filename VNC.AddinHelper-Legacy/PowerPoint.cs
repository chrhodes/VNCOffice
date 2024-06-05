namespace VNC.AddinHelper
{
    public class PowerPoint
    {
        private Microsoft.Office.Interop.PowerPoint.Application _PowerPointApplication;
        public Microsoft.Office.Interop.PowerPoint.Application PowerPointApplication
        {
	        get { return _PowerPointApplication; }
	        set { _PowerPointApplication = value; }
        }


        private bool _enableScreenUpdatesToggle = true;
        public bool EnableScreenUpdatesToggle
        {
	        get { return _enableScreenUpdatesToggle; }
	        set { _enableScreenUpdatesToggle = value; }
        }

    }
}
