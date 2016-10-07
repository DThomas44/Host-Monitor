/*
    flushdns.ahk
    Author: Grendahl

    Flushes the computer's DNS cache.
*/
if WinExist("Clear DNS Cache")
   ExitApp
else
   Run, %ComSpec% /c title Clear DNS Cache && ipconfig /flushdns && PAUSE
ExitApp
