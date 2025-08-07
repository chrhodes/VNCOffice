
$folders = @(
    "VNC.Visio.VSTOAddIn"
    , "VNC.VSTOAddin"
    , "VNCVisioTools"
    , "VNCVisioToolsApplication"
    , "VNCVisioToolsApplication.Core"    
    )

foreach ($folder in $folders)
{
    "Removing obj\ and bin\ folder contents in $folder"

    if (Test-Path -Path $folder\obj)
    {
        remove-item $folder\obj -Recurse -Force
    }

    if (Test-Path -Path $folder\bin)
    {
        remove-item $folder\bin -Recurse -Force
    }

}

Read-Host -Prompt "Press Enter to Exit"