using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using VNC.Core.Presentation;
using VNC.WPF.Presentation.Dx.Views;

using VNCExcelToolsApplication.Presentation.ViewModels;
using VNCExcelToolsApplication.Presentation.Views;

namespace VNCExcelToolsApplication.Actions
{
    public class Excel_CustomUI
    {
        public static DxThemedWindowHost folderMapHost = null;

        public static void FolderMap()
        {
            if (folderMapHost is null) folderMapHost = new DxThemedWindowHost(Common.EventAggregator);

            folderMapHost.DisplayUserControlInHost(
                "Folder Map",
                Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
                ShowWindowMode.Modeless_Show,
                //(CommandCockpitViewModel)Common.ApplicationBootstrapper.Container.Resolve(typeof(CommandCockpitViewModel))
                //(FolderMap)Common.ApplicationBootstrapper.Container.Resolve(typeof(FolderMap))
                new FolderMap()
            );
        }      
    }
}
