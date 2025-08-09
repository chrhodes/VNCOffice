using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Tools.Ribbon;

using Visio = Microsoft.Office.Interop.Visio;

namespace VNCShapeSheet
{
    public partial class Ribbon
    {
        // NOTE(crhodes)
        // This was moved out of designer so we can log

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();

            Int64 startTicks = Common.WriteToDebugWindow("Ribbon()", true);
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Int64 startTicks = Common.WriteToDebugWindow("Ribbon_Load()", true);
        }


    }
}
