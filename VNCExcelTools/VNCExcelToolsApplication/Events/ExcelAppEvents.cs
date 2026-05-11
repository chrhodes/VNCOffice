using System;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

namespace VNCExcelToolsApplication.Events
{
    public class ExcelAppEvents
    {
        private Application _ExcelApplication;

        public Application ExcelApplication
        {
            get
            {
                return _ExcelApplication;
            }
            set
            {
                if (_ExcelApplication != null)
                {
                    // Should remove all the event handlers;
                }

                _ExcelApplication = value;

                if (_ExcelApplication != null)
                {

                }
            }
        }

        #region Regular Events      

        // short _count_AfterModal;
        // void VisioApplication_AfterModal(Application app)
        // {
        // DisplayEventInWatchWindow(_count_AfterModal++, MethodInfo.GetCurrentMethod().Name);
        // }

        #endregion

        #region Chatty Events - Log if DisplayChattyEvents

        // short _count_CellChanged;
        // void VisioApplication_CellChanged(Cell Cell)
        // {
        // DisplayEventInWatchWindow(_count_CellChanged++, MethodInfo.GetCurrentMethod().Name, true);
        // }


        #endregion Chatty Events

        internal void DisplayEventInWatchWindow(short i, string outputLine, Boolean isChattyEvent = false)
        {
            if (Common.DisplayEvents)
            {
                if (isChattyEvent && !Common.DisplayChattyEvents) return;

                Common.WriteToWatchWindow($"{outputLine}:{i}");
            }
        }
    }
}
