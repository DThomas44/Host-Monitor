/*
    PingMsg.ahk
    Author: Daniel Thomas

    This script takes a target script name and a host name or IP address as
    parameters and sends a system message to the target script with the reply
    time of the provided host name or IP address. Returns TIMEOUT if no reply.
*/
;<=====  System Settings  =====================================================>
#SingleInstance Off
#NoEnv
#NoTrayIcon
DllCall("AllocConsole")
WinHide % "ahk_id " DllCall("GetConsoleWindow", "ptr")

;<=====  Parameters  ==========================================================>
targetScript = %1%
host = %2%
hostID = %3%

;<=====  Main  ================================================================>
ping := CheckServer(host)
ping := strSplit(ping, "|")
if ((ping[2] == -1)||ping[2] == ""){
    reply := hostID . "||TIMEOUT"
} else {
    reply := hostID . "|" . ping[1] . "|" . ping[2] . "|" . ping[3]
}
Send_WM_COPYDATA(reply, targetScript . " ahk_class AutoHotkey")
ExitApp

;<=====  Functions  ===========================================================>
CheckServer(host){
    objShell := ComObjCreate("WScript.Shell")
    objExec := objShell.Exec(ComSpec . " /c ping -a -n 1 " . host)
    strStdOut := ""
    while, !objExec.StdOut.AtEndOfStream
        strStdOut := objExec.StdOut.ReadAll()
    Loop, Parse, strStdOut, `n, `r
    {
        if (A_Index == 2)
        {
            str := strSplit(A_LoopField, A_Space)
            if (str[2] == host)
                str[2] := "<No DNS Resolution>"
            else
            {
                RegExMatch(A_LoopField, "\[((([1-9])|([1-9][0-9])|(1[0-9]{2})|([1-2][0-5]{2}))(\.(([0-9])|([1-9][0-9])|(1[0-9]{2})|([1-2][0-5]{2}))){3})\]", ip)
                StringTrimLeft, ip, ip, 1
                StringTrimRight, ip, ip, 1
            }
        }
        if (A_Index == 3)
        {
            RegExMatch(A_LoopField, "O)(?=|<)\d* ?ms", time)
        }
        else
            continue
    }
    return str[2] . "|" . time.Value() . "|" . ip
}

Send_WM_COPYDATA(ByRef StringToSend, ByRef TargetScriptTitle){
    VarSetCapacity(CopyDataStruct, 3*A_PtrSize, 0)
    SizeInBytes := (StrLen(StringToSend) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(SizeInBytes, CopyDataStruct, A_PtrSize)
    NumPut(&StringToSend, CopyDataStruct, 2*A_PtrSize)
    Prev_DetectHiddenWindows := A_DetectHiddenWindows
    Prev_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
    SendMessage, 0x4a, 0, &CopyDataStruct,, %TargetScriptTitle%
    DetectHiddenWindows %Prev_DetectHiddenWindows%
    SetTitleMatchMode %Prev_TitleMatchMode%
    return ErrorLevel
}
