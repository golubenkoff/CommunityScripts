<#
.SYNOPSIS
    Launches an application and moves its window to the second screen.

.DESCRIPTION
    This script starts a specified application and moves its main window to the second screen if available.
    It uses Windows API calls to manipulate the window position.

.PARAMETER applicationName
    The name of the application to launch (default is 'notepad').

.EXAMPLE
    .\Start-AppOnSecondScreen.ps1 -applicationName "notepad"
    Launches Notepad and moves its window to the second screen.

.NOTES
    - Ensure the application specified has a GUI window.
    - Requires at least two monitors connected to the system.
    - The script uses Windows API calls via Add-Type.

.LINK
		Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)


#>



<#

    .NOTES
████████╗ █████╗ ██╗  ██╗███████╗    ██╗████████╗
╚══██╔══╝██╔══██╗██║ ██╔╝██╔════╝    ██║╚══██╔══╝
   ██║   ███████║█████╔╝ █████╗      ██║   ██║
   ██║   ██╔══██║██╔═██╗ ██╔══╝      ██║   ██║
   ██║   ██║  ██║██║  ██╗███████╗    ██║   ██║▄█╗
   ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝    ╚═╝   ╚═╝╚═╝

██╗  ██╗ ██████╗ ██╗     ██████╗     ██╗████████╗
██║  ██║██╔═══██╗██║     ██╔══██╗    ██║╚══██╔══╝
███████║██║   ██║██║     ██║  ██║    ██║   ██║
██╔══██║██║   ██║██║     ██║  ██║    ██║   ██║
██║  ██║╚██████╔╝███████╗██████╔╝    ██║   ██║▄█╗
╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═════╝     ╚═╝   ╚═╝╚═╝

██╗      ██████╗ ██╗   ██╗███████╗    ██╗████████╗
██║     ██╔═══██╗██║   ██║██╔════╝    ██║╚══██╔══╝
██║     ██║   ██║██║   ██║█████╗      ██║   ██║
██║     ██║   ██║╚██╗ ██╔╝██╔══╝      ██║   ██║
███████╗╚██████╔╝ ╚████╔╝ ███████╗    ██║   ██║
╚══════╝ ╚═════╝   ╚═══╝  ╚══════╝    ╚═╝   ╚═╝

#>


param(
    $applicationName = 'notepad' # Default application to launch
)

# Create a new process for the application
$startup = Get-WmiObject Win32_ProcessStartup
$arguments = @{
    CommandLine      = $applicationName
    CurrentDirectory = 'c:\windows\system32'
}
$NewProcessID = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments $arguments

# Load required .NET assemblies for screen and window manipulation
[void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')
[void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

# Add Windows API functions for window manipulation
Add-Type -TypeDefinition '
using System;
using System.Runtime.InteropServices;

namespace AMEE {
    public static partial class WindowsAPI
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool SetWindowText(IntPtr hwnd, String lpString);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        public struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }
}
'

# Get all connected screens
$Screens = [System.Windows.Forms.Screen]::AllScreens
if ($Screens.Count -lt 2) {
    Write-Host 'Second screen not detected!'
    exit
}

# Get the bounds of the second screen
$SecondScreen = $Screens[1]
$X = $SecondScreen.Bounds.X
$Y = $SecondScreen.Bounds.Y
$Width = $SecondScreen.Bounds.Width
$Height = $SecondScreen.Bounds.Height

# Wait for the application to start and get its process
$ProcessList = $null
do {
    $ProcessList = Get-Process $applicationName | Where-Object Id -EQ $NewProcessID.ProcessId | Where-Object MainWindowTitle
} while ($null -eq $ProcessList)

# Move the application window to the second screen
foreach ($Process in $($ProcessList | Where-Object { $_.mainWindowTitle })) {
    $Owner = Get-CimInstance Win32_Process -Filter "ProcessId = '$($Process.Id)'" | ForEach-Object { Invoke-CimMethod -InputObject $_ -MethodName GetOwner }
    [AMEE.WindowsAPI]::MoveWindow($Process.mainWindowHandle, $X + 1, $Y + 1, 800, 600, $true)
    Write-Warning "Moving: [$applicationName] to Screen: 2"
}