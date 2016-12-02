/*
    Shares.ahk
    Author: Grendahl

    This script takes  host name, host login name, and host password as
    parameters and provides the user with a list of shares on the host.
    The user can then select a share and open it using the passed in credentials.

    Changes by Elesar:
        Added splitting of domain and username
        Added RunAs so you don't get explorer prompting for user/password (currently not working)

*/
;<=====  System Settings  =====================================================>
#SingleInstance Off
#NoEnv
#NoTrayIcon

;<=====  User Input  ==========================================================>
Host = %1%
User = %2%
Pass = %3%
If(!Host || !User || !Pass) {
    Msgbox Invalid Input - Must have Host/User/Pass
    ExitApp
}

if InStr(User, "\") {
    Domain := subStr(User, 1, inStr(User, "\") - 1)
    User := subStr(User, inStr(User, "\") + 1)
}

;<=====  Startup  =============================================================>
;Enumerate shares on host
try {
    Drives := ListShares(Host, User, Pass)
} catch e {
    MsgBox, % "Failed to get drives on " . host
    ExitApp
}

;<=====  GUI  =================================================================>
Menu, Tray, Icon, % A_ScriptDir . "\..\img\Host Monitor.ico"
Gui, Add, Text,,Choose a share
Gui, Add, ListBox, vShareSelection r%Count%, % Drives
Gui, Add, Button, gButton, GO!
Gui, Submit
Gui, Show
return

;<=====  Labels  ==============================================================>
Button:
    GuiControlGet, Share,,ListBox1
    Share := "\\" . Host . "\" . Share
    Gui, Destroy
    ;RunAs, %User%, %Pass%, %Domain%
    Run, Explorer /n`,/e`,%Share%
ExitApp

GuiClose:
    ExitApp

ESC::ExitApp

;<=====  Functions  ===========================================================>
ListShares(Server, User, Pass) {
    PropertyList := "Name"
    wmiLocator := ComObjCreate("WbemScripting.SWbemLocator")
    objWMIService := wmiLocator.ConnectServer(Server, "root\cimv2", User, Pass)
    WQLQuery = Select * From Win32_Share
    colDiskDrive := objWMIService.ExecQuery(WQLQuery)._NewEnum
    While colDiskDrive[objDiskDrive]
        Loop, Parse, PropertyList, `,
        {
            Global Count ++
            Shares  .= A_index = 1 ? objDiskDrive[A_LoopField] . "|"  :  . A_LoopField
        }
    StringTrimRight, Shares, Shares, 1
    Return, Shares
}
