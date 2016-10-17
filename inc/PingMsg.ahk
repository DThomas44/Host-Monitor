/*
    PingMsg.ahk
    Author: Daniel Thomas

    This script takes a target script name host name or IP address and hostID as
    parameters and sends a system message to the target script with the reply
    time of the provided host name or IP address. Returns TIMEOUT if no reply.

    Reply is in the format of: hostID|hostName|latency|IPv4Address

    Hostname is only provided if it was used as a parameter. Currently no
    reverse lookup available.
*/
;<=====  System Settings  =====================================================>
#SingleInstance Off
#NoEnv
#NoTrayIcon

;<=====  Parameters  ==========================================================>
targetScript = %1%
host = %2%
hostID = %3%

;<=====  Main  ================================================================>
reply := hostID . "|"

if RegExMatch(host, "^((|\.)\d{1,3}){4}$")
{
    try {
        reply .= IPHelper.ReverseLookup(host) . "|"
    }
    catch e {
        reply .= "|"
    }
}
else
    reply .= host . "|"

try {
    reply .= IPHelper.Ping(host) . "|"
}
catch e {
    reply .= "TIMEOUT|"
}

if !(RegExMatch(addr, "^((|\.)\d{1,3}){4}$"))
{
    Try {
        reply .= IPHelper.ResolveHostname(host)
    }
    catch e {
        reply .= "|"
    }
}
else
    reply .= host

Send_WM_COPYDATA(reply, targetScript . " ahk_class AutoHotkey")
ExitApp

;<=====  Functions  ===========================================================>
Send_WM_COPYDATA(ByRef StringToSend, ByRef TargetScriptTitle){
    VarSetCapacity(CopyDataStruct, 3*A_PtrSize, 0)
    SizeInBytes := (StrLen(StringToSend) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(SizeInBytes, CopyDataStruct, A_PtrSize)
    NumPut(&StringToSend, CopyDataStruct, 2*A_PtrSize)
    Prev_DetectHiddenWindows := A_DetectHiddenWindows
    Prev_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
    SendMessage, 0x4a, 0, &CopyDataStruct,, %TargetScriptTitle%,,,, 10000
    DetectHiddenWindows %Prev_DetectHiddenWindows%
    SetTitleMatchMode %Prev_TitleMatchMode%
    return ErrorLevel
}

;<=====  Includes  ============================================================>
#Include %A_ScriptDir%\IPHelper.ahk
