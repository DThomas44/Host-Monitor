/*
    pingGraph.ahk
    Author: Elesar (Daniel Thomas)
    Based mostly on demo script for XGraph.ahk by GeekDude

    Pings a hostname/IP of computer provided as first arg. constantly and plots
        results as graph.

    To Do:
        Implement logging
*/
;<=====  System Settings  =====================================================>
#SingleInstance, Force
#NoEnv
ListLines, Off

;<=====  Startup  =============================================================>
;Receive messages from slave scripts
OnMessage(0x4a, "Receive_WM_COPYDATA")

;Allcate console so we don't have constant flashing windows while pinging
DllCall("AllocConsole")
WinHide % "ahk_id " DllCall("GetConsoleWindow", "ptr")

;<=====  User Input  ==========================================================>
host = %1%

;<=====  GUI  =================================================================>
Menu, Tray, Icon, % A_ScriptDir . "\..\img\Host Monitor.ico"

Gui, Margin, 5, 5
Gui, Font, s8, Verdana
Gui, Add, CheckBox, x5 y5 w100 vLogResults gLogResults, Log Results
Gui, Add, Text, x+5 yp w300 vLogFileText,
Loop, % 11 + ( Y := 30 ) - 30 ; Loop 11 times
   Gui, Add, Text, xm y%y% w22 h10 0x200 Right, % 140 - (Y += 10)
ColumnW := 10
hBM := XGraph_MakeGrid(  ColumnW, 10, 40, 12, 0x008000, 0, GraphW, GraphH )
Gui, Add, Text, % "xm+25 ym+20 w" ( GraphW + 2 ) " h" ( GraphH + 2 ) " 0x1000"
Gui, Add, Text, xp+1 yp+1 w%GraphW% h%GraphH% hwndhGraph1 0xE, pGraph1
pGraph1 := XGraph( hGraph1, hBM, ColumnW, "1,10,0,10", 0x00FF00, 1, True )
Gui, Show,, XGraph Ping %host%
Loop, 100
    XGraph_Plot(pGraph1, A_Index)
SetTimer, Ping, 1000
Return

;<=====  Labels  ==============================================================>
GuiClose:
    ExitApp

LogResults:
    Gui, Submit, NoHide
    if LogResults
    {
        logFile := A_ScriptDir . "\..\Logs\" . host . "_" . A_DD . A_MM . A_Hour . A_Min . ".txt"
        GuiControl,, LogFileText, % logFile
    }
    else
        GuiControl,, LogFileText,
    Return

Ping:
    Run, % A_ScriptDir . "\PingMsg.ahk """ . A_ScriptName . """ "
        . host . " " . A_Index,, Hide, threadID
    Return

;<=====  Functions  ===========================================================>
Receive_WM_COPYDATA(wParam, lParam){
    Global
    StringAddress := NumGet(lParam + 2*A_PtrSize)
    CopyOfData := StrGet(StringAddress)
    reply := StrSplit(CopyOfData, "|")
    if (reply[3] != "TIMEOUT")
    {
        ;Good return
        XGraph_Plot(pGraph1, 100 - strReplace(reply[3], "ms"))
    } else {
        ;Timeout
        XGraph_Plot(pGraph1, 9999)
    }
    return true
}

;<=====  Includes  ============================================================>
#Include %A_ScriptDir%
#Include Common Functions.ahk
#Include XGraph.ahk
