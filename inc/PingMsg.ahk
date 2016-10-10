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
result := Object()
if !Ping4(host, result, 1500) {
    reply := hostID . "||TIMEOUT"
} else {
    reply := hostID . "|" . result.Name . "|" . result.RTTime . "ms|" . result.IPAddr
}

Send_WM_COPYDATA(reply, targetScript . " ahk_class AutoHotkey")
ExitApp

;<=====  Functions  ===========================================================>

; ======================================================================================================================
; Function:       IPv4 ping with name resolution, based upon 'SimplePing' by Uberi ->
;                 http://www.autohotkey.com/board/topic/87742-simpleping-successor-of-ping/
; Parameters:     Addr     -  IPv4 address or host / domain name.
;                 ----------  Optional:
;                 Result   -  Object to receive the result in three keys:
;                             -  InAddr - Original value passed in parameter Addr.
;                             -  IPAddr - The replying IPv4 address.
;                             -  Name   - The replying host name.
;                             -  RTTime - The round trip time, in milliseconds.
;                 Timeout  -  The time, in milliseconds, to wait for replies.
; Return values:  On success: The round trip time, in milliseconds.
;                 On failure: "", ErrorLevel contains additional informations.
; AHK version:    AHK 1.1.13.01
; Tested on:      Win XP - AHK A32/U32, Win 7 x64 - AHK A32/U32/U64
; Authors:        Uberi / just me
; Version:        1.0.00.00/2013-11-06/just me - initial release
;                 1.0.01.00/2013-12-01/just me - added return of host name
; MSDN:           Winsock Functions   -> http://msdn.microsoft.com/en-us/library/ms741394(v=vs.85).aspx
;                 IP Helper Functions -> hhttp://msdn.microsoft.com/en-us/library/aa366071(v=vs.85).aspx
; ======================================================================================================================
Ping4(Addr, ByRef Result := "", Timeout := 1024) {
    ; ICMP status codes -> http://msdn.microsoft.com/en-us/library/aa366053(v=vs.85).aspx
    ; WSA error codes   -> http://msdn.microsoft.com/en-us/library/ms740668(v=vs.85).aspx
    Static WSADATAsize := (2 * 2) + 257 + 129 + (2 * 2) + (A_PtrSize - 2) + A_PtrSize
    OrgAddr := Addr
    Result := ""
    ; -------------------------------------------------------------------------------------------------------------------
    ; Initiate the use of the Winsock 2 DLL
    VarSetCapacity(WSADATA, WSADATAsize, 0)
    If (Err := DllCall("Ws2_32.dll\WSAStartup", "UShort", 0x0202, "Ptr", &WSADATA, "Int")) {
        ErrorLevel := "WSAStartup failed with error " . Err
        Return ""
    }
    If !RegExMatch(Addr, "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$") { ; Addr contains a name
        If !(HOSTENT := DllCall("Ws2_32.dll\gethostbyname", "AStr", Addr, "UPtr")) {
            DllCall("Ws2_32.dll\WSACleanup") ; Terminate the use of the Winsock 2 DLL
            ErrorLevel := "gethostbyname failed with error " . DllCall("Ws2_32.dll\WSAGetLastError", "Int")
            Return ""
        }
        PAddrList := NumGet(HOSTENT + 0, (2 * A_PtrSize) + 4 + (A_PtrSize - 4), "UPtr")
        PIPAddr   := NumGet(PAddrList + 0, 0, "UPtr")
        Addr := StrGet(DllCall("Ws2_32.dll\inet_ntoa", "UInt", NumGet(PIPAddr + 0, 0, "UInt"), "UPtr"), "CP0")
    }
    INADDR := DllCall("Ws2_32.dll\inet_addr", "AStr", Addr, "UInt") ; convert address to 32-bit UInt
    If (INADDR = 0xFFFFFFFF) {
        ErrorLevel := "inet_addr failed for address " . Addr
        DllCall("Ws2_32.dll\WSACleanup")
        Return ""
    }
    ; -------------------------------------------------------------------------------------------------------------------
    HMOD := DllCall("LoadLibrary", "Str", "Iphlpapi.dll", "UPtr")
    Err := ""
    If (HPORT := DllCall("Iphlpapi.dll\IcmpCreateFile", "UPtr")) { ; open a port
        REPLYsize := 32 + 8
        VarSetCapacity(REPLY, REPLYsize, 0)
        If DllCall("Iphlpapi.dll\IcmpSendEcho", "Ptr", HPORT, "UInt", INADDR, "Ptr", 0, "UShort", 0
                                              , "Ptr", 0, "Ptr", &REPLY, "UInt", REPLYsize, "UInt", Timeout, "UInt") {
            Result := {}
            Result.InAddr := OrgAddr
            Result.IPAddr := StrGet(DllCall("Ws2_32.dll\inet_ntoa", "UInt", NumGet(Reply, 0, "UInt"), "UPtr"), "CP0")
            Result.RTTime := NumGet(Reply, 8, "UInt")
        }
        Else
            Err := "IcmpSendEcho failed with error " . A_LastError
        DllCall("Iphlpapi.dll\IcmpCloseHandle", "Ptr", HPORT)
    }
    Else
        Err := "IcmpCreateFile failed to open a port!"
    DllCall("FreeLibrary", "Ptr", HMOD)
    ; -------------------------------------------------------------------------------------------------------------------
    If (Err) {
        DllCall("Ws2_32.dll\WSACleanup")
        ErrorLevel := Err
        Return ""
    }
    VarSetCapacity(INADDR, 4, 0)
    NumPut(DllCall("Ws2_32.dll\inet_addr", "AStr", Result.IPAddr, "UInt"), INADDR, 0, "UInt")
    If (HOSTENT := DllCall("Ws2_32.dll\gethostbyaddr", "Ptr", &INADDR, "Int", 4, "Int", 2, "UPtr"))
        Result.Name := StrGet(NumGet(HOSTENT + 0, 0, "UPtr"), 256, "CP0")
    Else
        Result.Name := ""
    DllCall("Ws2_32.dll\WSACleanup")
    ErrorLevel := 0
    Return Result.RTTime
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
