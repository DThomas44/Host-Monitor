/*
    Host-Monitor Build Script
    Author: Daniel Thomas
*/
;<=====  System Settings  =====================================================>
#SingleInstance Force
#NoEnv

;<=====  Clean up bin  ========================================================>
IfExist, % A_ScriptDir . "\bin"
{
    try {
        FileRemoveDir, % A_ScriptDir . "\bin", 1
    } catch {
        MsgBox, % "Could not delete old bin folder.`nAborting compile."
        ExitApp
    }
}
FileCreateDir, % A_ScriptDir . "\bin"
FileCreateDir, % A_ScriptDir . "\temp"

;<=====  Compile tools  =======================================================>
RunWait, %ComSpec% /c ""C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe" /in "%A_ScriptDir%\inc\pingGraph.ahk" /out "%A_ScriptDir%\temp\pingGraph.exe" >"%A_ScriptDir%\CompileResults.txt""
RunWait, %ComSpec% /c ""C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe" /in "%A_ScriptDir%\inc\pingMsg.ahk" /out "%A_ScriptDir%\temp\pingMsg.exe" >>"%A_ScriptDir%\CompileResults.txt""
RunWait, %ComSpec% /c ""C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe" /in "%A_ScriptDir%\inc\traceRt.ahk" /out "%A_ScriptDir%\temp\traceRt.exe" >>"%A_ScriptDir%\CompileResults.txt""

;<=====  Compile main  ========================================================>
RunWait, %ComSpec% /c ""C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe" /in "%A_ScriptDir%\Host Monitor.ahk" /out "%A_ScriptDir%\bin\Host Monitor.exe" /icon "%A_ScriptDir%\img\Host Monitor.ico" >>"%A_ScriptDir%\CompileResults.txt""

;<=====  Cleanup  =============================================================>
try {
    FileRemoveDir, % A_ScriptDir . "\temp", 1
} catch {
    MsgBox, % "Could not delete temp folder."
}

;<=====  Finished  ============================================================>
MsgBox, Done.
ExitApp
