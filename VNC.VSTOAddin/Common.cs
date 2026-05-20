using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace VNC.VSTOAddIn
{
    public class Common : VNC.Core.Common
    {
        public new const string LOG_CATEGORY = "VSTOAddIn";

        public static Boolean EnableAppEvents = false;  // Custom Header and Footer need this enabled.
        public static Boolean DisplayEvents = false;
        public static Boolean DisplayChattyEvents = false;

        public static Boolean EnableLogging
        {
            get;
            set;
        }

        private static Presentation.frmDebugWindow _DebugWindow;
        public static Presentation.frmDebugWindow DebugWindow
        {
            get
            {
                if (_DebugWindow == null)
                {
                    _DebugWindow = new Presentation.frmDebugWindow();
                }

                return _DebugWindow;
            }
            set
            {
                _DebugWindow = value;
            }
        }

        private static Presentation.frmWatchWindow _WatchWindow;
        public static Presentation.frmWatchWindow WatchWindow
        {
            get
            {
                if (_WatchWindow == null)
                {
                    _WatchWindow = new Presentation.frmWatchWindow();
                }
                return _WatchWindow;
            }
            set
            {
                _WatchWindow = value;
            }
        }

        // TODO(crhodes)
        // Based on how these are called, seems like the WatchWindow is for "watching" the XlLocation stuff
        // and the DebugWindow is for general debugging.
        // Looks like we started to use CallerMemberName to identify the calling member in the WatchWindow
        // but not in the DebugWindow.
        // We should probably be consistent and use CallerMemberName in both places.
        public static long WriteToWatchWindow(string message, [CallerMemberName] string callingMember = "")
        {
            if (DeveloperMode)
            {
                WatchWindow.AddOutputLine($"{callingMember}: {message}");
            }

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToWatchWindow(string message, long startTicks, [CallerMemberName] string callingMember = "")
        {
            if (DeveloperMode)
            {
                WatchWindow.AddOutputLine($"{callingMember}: {message}"
                    + "-" + (double)(Stopwatch.GetTimestamp() - startTicks) / (double)Stopwatch.Frequency);
            }

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToDebugWindow(string message, Boolean force = false)
        {
            if (DeveloperMode || force)
            {
                DebugWindow.AddOutputLine(message);
            }

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToDebugWindow(string message, long startTicks, Boolean force = false)
        {
            if (DeveloperMode || force)
            {
                DebugWindow.AddOutputLine(message + "-" + (double)(Stopwatch.GetTimestamp() - startTicks) / (double)Stopwatch.Frequency);
            }

            return Stopwatch.GetTimestamp();
        }
    }
}
