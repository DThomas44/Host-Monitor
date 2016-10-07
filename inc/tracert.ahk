/*
    tracert.ahk
    Author: Elesar (Daniel Thomas)

    Starts a tracert to hostname/IP provided as first arg.
*/
if WinExist("Trace " 1)
   ExitApp
else
   Run, %ComSpec% /c title Trace %1% && tracert %1% && PAUSE
ExitApp
