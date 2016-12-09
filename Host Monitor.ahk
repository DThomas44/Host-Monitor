/*
    Host Monitor.ahk
    Authors: Elesar (Daniel Thomas) & Grendahl

    Uses functions by SKAN, just me, Uberi & jNizM

    Monitors a set of hosts periodically and gives visual representation of their status.
    Includes several tools for troubleshooting.

    Some functionality depends on putty.exe being in the inc folder.

*/
;<=====  System Settings  =====================================================>
#SingleInstance, Force
#NoEnv
SetBatchLines, -1

;<=====  FileInstalls  ========================================================>
;Images
IfNotExist, % A_ScriptDir . "\img"
    FileCreateDir, % A_ScriptDir . "\img"
;FileInstall, img\Green.jpg, img\Green.jpg, 1
;FileInstall, img\Yellow.jpg, img\Yellow.jpg, 1
;FileInstall, img\Red.jpg, img\Red.jpg, 1
FileInstall, img\Host Monitor.ico, img\Host Monitor.ico, 1

;Tools
IfNotExist, % A_ScriptDir . "\inc"
    FileCreateDir, % A_ScriptDir . "\inc"
FileInstall, temp\pingGraph.exe, inc\pingGraph.exe, 1
FileInstall, temp\pingMsg.exe, inc\pingMsg.exe, 1
FileInstall, temp\traceRt.exe, inc\traceRt.exe, 1

;Settings & Misc
FileInstall, inc\defaultSettings.xml, inc\settings.xml, 0
FileInstall, inc\transform.xslt, inc\transform.xslt, 1
FileInstall, hosts.txt, hosts.txt, 0

;Clean up file blocking
Loop, %A_WorkingDir%\*.*, 0, 1 ;Loop all files from A_WorkingDir and subdirectories
    FileDelete, fileName:Zone.Identifier$DATA ;Remove Zone Identifier Data to Unblock file

;<=====  Startup  =============================================================>
;Receive messages from slave scripts and Windows
OnMessage(0x4a, "Receive_WM_COPYDATA")
OnMessage(0x200, "WM_MOUSEMOVE") ; Tooltips

;Allcate console so we don't have constant flashing windows while pinging
DllCall("AllocConsole")
WinHide % "ahk_id " DllCall("GetConsoleWindow", "ptr")

;Make sure we are exiting nicely
OnExit, Exit

;<=====  Settings  ============================================================>
;Static Settings - May mess stuff up visually if changed
boxWidth := 150
boxHeight := 45

;Read in settings.xml
settings := loadXML(A_ScriptDir . "\inc\settings.xml")

;Read in transform.xslt
tdoc := loadXML(A_ScriptDir . "\inc\transform.xslt")

;Get monitor working area
SysGet, Mon, MonitorWorkArea
maxRows := Floor(MonBottom / boxHeight) - 2

;Adjust guiRows setting if higher than maxRows
if (settings.selectSingleNode("/hostMonitor/settings/guiRows").text > maxRows) {
    node := settings.selectSingleNode("/hostMonitor/settings/guiRows")
    node.text := maxRows
    SaveSettings(settings, tdoc)
}

;<=====  Read in hostList  ====================================================>
;Prompt for host file if not saved.
while !settings.selectSingleNode("/hostMonitor/settings/hostPath").text
{
    FileSelectFile, path,,, Select host file, Text (*.txt)
    if !path
    {
        MsgBox, 4,, No file selected.`nDo you want to exit?
        IfMsgBox, Yes
            ExitApp
    } else {
        node := settings.selectSingleNode("/hostMonitor/settings/hostPath")
        node.text := path
        SaveSettings(settings, tdoc)
    }
}

hostsFile := fileOpen(settings.selectSingleNode("/hostMonitor/settings/hostPath").text, "r")
hostsData := hostsFile.Read()
hostsFile.Close()

hosts := Object()
Loop, Parse, hostsData, `n, `r
{
    if (A_LoopField != "")
    {
        hosts[A_Index] := Object()
        aliasPos := RegExMatch(A_LoopField, "\[(.*?)\]", alias)
        alias := strReplace(alias, "[")
        alias := strReplace(alias, "]")
        if alias
        {
            hosts[A_Index, "name"] := subStr(A_LoopField, 1, aliasPos - 1)
            hosts[A_Index, "alias"] := alias
        }
        else
        hosts[A_Index, "name"] := A_LoopField
        hosts[A_Index, "ip"] := ""
        hosts[A_Index, "lastSeen"] := "Never"
        hosts[A_Index, "bgImageID"] := ""
        hosts[A_Index, "nameTextID"] := ""
        hosts[A_Index, "statusTextID"] := ""
        hosts[A_Index, "threadID"] := ""
        hosts[A_Index].pingArray := Object()
    }
}

;<=====  Menus  ===============================================================>
;Context Menu
Menu, ContextMenu, Add, Trace Route, ContextMenuHandler
Menu, ContextMenu, Add, Ping Graph, ContextMenuHandler
Menu, ContextMenu, Add
Menu, ContextMenu, Add, RDP, ContextMenuHandler
Menu, ContextMenu, Add
Menu, ContextMenu, Add, Explore Shares, ContextMenuHandler
Menu, ContextMenu, Add
Menu, ContextMenu, Add, Browse, ContextMenuHandler
Menu, ContextMenu, Add, Browse (https), ContextMenuHandler
Menu, ContextMenu, Add
Menu, ContextMenu, Add, SSH, ContextMenuHandler
Menu, ContextMenu, Add, Telnet, ContextMenuHandler

;File Menu
Menu, FileMenu, Add, &Scan Now, MenuHandler
Menu, FileMenu, Add, Open &Log Folder, MenuHandler
Menu, FileMenu, Add
Menu, FileMenu, Add, &Open, MenuHandler
Menu, FileMenu, Add, &Reload, MenuHandler
Menu, FileMenu, Add
Menu, FileMenu, Add, E&xit, MenuHandler

;Help Menu
Menu, HelpMenu, Add, &Debug, MenuHandler
Menu, HelpMenu, Add
Menu, HelpMenu, Add, &About, MenuHandler

;Settings Menu
Menu, SettingsMenu, Add, &Flush DNS, SettingsMenuHandler
Menu, SettingsMenu, Add
Menu, SettingsMenu, Add, &Logging, SettingsMenuHandler
if settings.selectSingleNode("/hostMonitor/settings/logging").text
    Menu, SettingsMenu, Check, &Logging
Menu, SettingsMenu, Add, Auto &TraceRt, SettingsMenuHandler
if settings.selectSingleNode("/hostMonitor/settings/autoTraceRt").text
    Menu, SettingsMenu, Check, Auto &TraceRt
Menu, SettingsMenu, Add, Remember &Window Pos, SettingsMenuHandler
if settings.selectSingleNode("/hostMonitor/settings/rememberPos").text
    Menu, SettingsMenu, Check, Remember &Window Pos
Menu, SettingsMenu, Add, Pass &Credentials to Shares, SettingsMenuHandler
if settings.selectSingleNode("/hostMonitor/settings/useCreds").text
    Menu, SettingsMenu, Check, Pass &Credentials to Shares
Menu, SettingsMenu, Add
Menu, SettingsMenu, Add, Change GUI &Rows, SettingsMenuHandler
Menu, SettingsMenu, Add, Change Max &Threads, SettingsMenuHandler
Menu, SettingsMenu, Add, Change Check &Interval, SettingsMenuHandler
Menu, SettingsMenu, Add, Change Ping &Average Count, SettingsMenuHandler
Menu, SettingsMenu, Add, Change &Warning Latency, SettingsMenuHandler

;Menu Bar
Menu, MenuBar, Add, &File, :FileMenu
Menu, MenuBar, Add, &Settings, :SettingsMenu
Menu, MenuBar, Add, &Help, :HelpMenu

;Tray Menu
Menu, Tray, Icon, % A_ScriptDir . "\img\Host Monitor.ico"

;<=====  Load Images  =========================================================>
;hGreen := LoadPicture(A_ScriptDir . "\img\Green.jpg")
;hYellow := LoadPicture(A_ScriptDir . "\img\Yellow.jpg")
;hRed := LoadPicture(A_ScriptDir . "\img\Red.jpg")

;<=====  GUI  =================================================================>
Gui, Margin, 5, 5
Gui, Menu, MenuBar
RowX := 5
RowY := 5
Loop % hosts.MaxIndex() {
    if (A_Index != 1) && !Mod(A_Index - 1, (settings.selectSingleNode("/hostMonitor/settings/guiRows").text
            ?settings.selectSingleNode("/hostMonitor/settings/guiRows").text
            :maxRows)) {
        RowX += 155
        RowY := 5
    }

    ;Gui, Add, Picture, % "x" RowX " y" RowY " w" boxWidth " 0x4000000 gDummyLabel vi"
    ;    . A_Index, % "HBITMAP:*" hYellow

    Gui, Add, Progress, % "x" RowX " y" RowY " w" boxWidth " h" boxHeight " vi"
        . A_Index " cYellow", 100

    if hosts[A_Index, "alias"] {
        Gui, Add, Text, % "xp+5 yp+5 w140 h15 gDummyLabel +0x200 +BackgroundTrans vt"
            . A_Index . " +Center" , % subStr(hosts[A_Index, "alias"], 1, 25)
    }
    else {
        Gui, Add, Text, % "xp+5 yp+5 w140 h15 gDummyLabel +0x200 +BackgroundTrans vt"
            . A_Index . " +Center", % subStr(hosts[A_Index, "name"], 1, 25)
    }
    Gui, Add, Text, % "xp yp+15 w140 h15 gDummyLabel +0x200 +BackgroundTrans vs"
        . A_Index . " +Center", % hosts[A_Index, "lastSeen"]
    hosts[A_Index, "bgImageID"] := "i" . A_Index
    hosts[A_Index, "nameTextID"] := "t" . A_Index
    hosts[A_Index, "statusTextID"] := "s" . A_Index
    RowY += boxHeight
}
Gui, Add, StatusBar
SB_SetParts(100,100)
if (RowX < 100) {
    if settings.selectSingleNode("/hostMonitor/settings/rememberPos").text {
        Gui, Show,  % "x" . settings.selectSingleNode("/hostMonitor/settings/winX").text
            . " y" . settings.selectSingleNode("/hostMonitor/settings/winY").text . " w320", Host Monitor
    }
    else
        Gui, Show,  % "w320", Host Monitor
}
else {
    if settings.selectSingleNode("/hostMonitor/settings/rememberPos").text {
        Gui, Show, % "x" . settings.selectSingleNode("/hostMonitor/settings/winX").text
            . " y" . settings.selectSingleNode("/hostMonitor/settings/winY").text, Host Monitor
    }
    else
        Gui, Show, Center, Host Monitor
}

;<=====  Timers  ==============================================================>
;Check hosts based off of settings.xml
SetTimer, CheckHosts, % (settings.selectSingleNode("/hostMonitor/settings/checkInterval").text * 1000)

;Update status bar every second
scanTimer := settings.selectSingleNode("/hostMonitor/settings/checkInterval").text
SetTimer, UpdateSB, 1000

;Clock time based events
SetTimer, CheckTime, 60000

;<=====  Main  ================================================================>
if settings.selectSingleNode("/hostMonitor/settings/logging").Text
    logFile := LogStart(settings)
CheckHosts()

;<=====  End AutoExecute  =====================================================>
return

;<=====  Subs  ================================================================>
About:
    aboutText =
    ( LTrim
        Written by Elesar (Daniel Thomas)
        In colaberation with Grendahl
        and others from the AutoHotkey community.
    )
    MsgBox, % aboutText
    return

CheckTime:
    ;Start a new log file if midnight
    if ((A_Hour == 00) && (A_Min == 00) && (settings.selectSingleNode("/hostMonitor/settings/logging").text))
    {
        logStop(logFile)
        logFile := logStart(settings)
    }
    return

ContextMenuHandler:
    if (A_ThisMenuItem == "Trace Route")
        StartTrace(host)
    else if (A_ThisMenuItem == "Ping Graph")
        StartPingGraph(host)
    else if (A_ThisMenuItem == "RDP")
        Run, % "mstsc /v:" . host
    else if (A_ThisMenuItem == "Explore Shares") {
        hostName := IPHelper.ReverseLookup(host)
        StringSplit, hostName, hostName,.
        Run, % A_ScriptDir "\inc\Shares.ahk " . hostName1
    }
    else if (A_ThisMenuItem == "Browse")
        Run, % "http://" . hosts[reply[1], "name"], % subStr(hosts[reply[1], "name"], 1, 25)
    else if (A_ThisMenuItem == "Browse (https)")
        Run, % "https://" . hosts[reply[1], "name"], % subStr(hosts[reply[1], "name"], 1, 25)
    else if (A_ThisMenuItem == "SSH")
        StartPutty(host, "ssh")
    else if (A_ThisMenuItem == "Telnet")
        StartPutty(host, "telnet")
    else
        MsgBox, % "Context menu broke :("
    return

Debug:
    clipboard := enumObj(hosts) . "`n`n" . enumObj(threads)
    MsgBox, % "Hosts and Threads arrays copied to clipboard."
    Return

DummyLabel:
    ;Needed to get tooltips/menus on text controls...
    Return

Exit:
GuiClose:
    WinGetPos, winX, winY,,, Host Monitor
    node := settings.selectSingleNode("/hostMonitor/settings/winX")
    node.text := winX
    node := settings.selectSingleNode("/hostMonitor/settings/winY")
    node.text := winY
    SaveSettings(settings, tdoc)
    if settings.selectSingleNode("/hostMonitor/settings/logging").text
        LogStop(logFile)
    ExitApp
    return

GuiContextMenu:
    host := hosts[subStr(A_GuiControl, 2), "ip"]
    if host
        Menu, ContextMenu, Show
    return

MenuButton:
    host := hosts[subStr(A_GuiControl, 2), "name"]
    Menu, ContextMenu, Show
    return

MenuHandler:
    if (A_ThisMenuItem == "&Scan Now")
        CheckHosts()
    else if (A_ThisMenuItem == "Open &Log Folder")
        Run, % "explore " . A_ScriptDir . "\Logs"
    else if (A_ThisMenuItem == "&Open"){
        FileSelectFile, path,,, Select host file, Text (*.txt)
        if path
        {
            node := settings.selectSingleNode("/hostMonitor/settings/hostPath")
            node.text := path
            SaveSettings(settings, tdoc)
            Reload
        }
    }
    else if (A_ThisMenuItem == "&Reload")
        Reload
    else if (A_ThisMenuItem == "E&xit")
        GoSub, Exit
    else if (A_ThisMenuItem == "&Debug")
        GoSub, Debug
    else if (A_ThisMenuItem == "&About")
        GoSub, About
    else
        MsgBox, % "Menu system broke :("
    return

SettingsMenuHandler:
    reloadScript := false
    if (A_ThisMenuItem == "&Flush DNS")
        MsgBox, % ((FlushDNS() == 1)?"DNS cache flushed.":"Failed to flush DNS cache.")
    else if (A_ThisMenuItem == "&Logging"){
        Menu, SettingsMenu, ToggleCheck, &Logging
        node := settings.selectSingleNode("/hostMonitor/settings/logging")
        node.text := !node.text
        if node.text
            logFile := LogStart(settings)
        else
            LogStop(logFile)
    }
    else if (A_ThisMenuItem == "Auto &TraceRt"){
        Menu, SettingsMenu, ToggleCheck, Auto &TraceRt
        node := settings.selectSingleNode("/hostMonitor/settings/autoTraceRt")
        node.text := !node.text
    }
    else if (A_ThisMenuItem == "Remember &Window Pos"){
        Menu, SettingsMenu, ToggleCheck, Remember &Window Pos
        node := settings.selectSingleNode("/hostMonitor/settings/rememberPos")
        node.text := !node.text
    }
    else if (A_ThisMenuItem == "Pass &Credentials to Shares"){
        Menu, SettingsMenu, ToggleCheck, Pass &Credentials to Shares
        node := settings.selectSingleNode("/hostMonitor/settings/useCreds")
        node.text := !node.text
    }
    else if (A_ThisMenuItem == "Change GUI &Rows"){
        InputBox, userInput, Change GUI Rows, % "Enter a row count between 1 and "
            . maxRows . ".`nSet to 0 for AutoSize based on monitor work area and host count."
            . "`nCurrent setting is " . settings.selectSingleNode("/hostMonitor/settings/guiRows").text
        if !userInput
            return
        if userInput is not integer
        {
            MsgBox, This setting can only take an integer.
            return
        }
        if (userInput > maxRows)
            userInput := maxRows
        node := settings.selectSingleNode("/hostMonitor/settings/guiRows")
        node.text := userInput
        MsgBox, 4,,% "Setting changed. This setting requires a script reload.`nReload now?"
        ifMsgBox, Yes
            reloadScript := true
    }
    else if (A_ThisMenuItem == "Change Max &Threads"){
        InputBox, userInput, Change Max Threads, % "Enter a thread count greater than 1."
            . "`nWarning: May load CPU or cause script issues if too many threads are allowed."
            . "`nCurrent setting is " . settings.selectSingleNode("/hostMonitor/settings/maxThreads").text
        if !userInput
            return
        if userInput is not integer
        {
            MsgBox, This setting can only take an integer.
            return
        }
        if (userInput < 1)
            userInput := 1
        node := settings.selectSingleNode("/hostMonitor/settings/maxThreads")
        node.text := userInput
    }
    else if (A_ThisMenuItem == "Change Check &Interval"){
        InputBox, userInput, Change Check Interval, % "Enter a check interval in seconds."
            . "`nSuggest 30 or higher."
            . "`nCurrent setting is " . settings.selectSingleNode("/hostMonitor/settings/checkInterval").text
        if !userInput
            return
        if userInput is not integer
        {
            MsgBox, This setting can only take an integer.
            return
        }
        if (userInput < 1)
            userInput := 1
        node := settings.selectSingleNode("/hostMonitor/settings/checkInterval")
        node.text := userInput
    }
    else if (A_ThisMenuItem == "Change Ping &Average Count"){
        InputBox, userInput, Change Ping Average Count, % "Enter an averaging count."
            . "`nSuggest 10 - 50."
            . "`nCurrent setting is " . settings.selectSingleNode("/hostMonitor/settings/pingAvgCount").text
        if !userInput
            return
        if userInput is not integer
        {
            MsgBox, This setting can only take an integer.
            return
        }
        if (userInput < 2)
            userInput := 2
        node := settings.selectSingleNode("/hostMonitor/settings/pingAvgCount")
        node.text := userInput
        MsgBox, 4,,% "Setting changed. This setting requires a script reload.`nReload now?"
        ifMsgBox, Yes
            reloadScript := true
    }
    else if (A_ThisMenuItem == "Change &Warning Latency"){
        InputBox, userInput, Change Ping Average Count, % "Enter a high latency warning level."
            . "`nCurrent setting is " . settings.selectSingleNode("/hostMonitor/settings/warnLatency").text
        if !userInput
            return
        if userInput is not integer
        {
            MsgBox, This setting can only take an integer.
            return
        }
        node := settings.selectSingleNode("/hostMonitor/settings/warnLatency")
        node.text := userInput
    }
    else
        MsgBox, % "Settings menu system broke :("
    SaveSettings(settings, tdoc)
    if reloadScript
        Reload
    return

UpdateSB:
    scanTimer--
    SB_SetText("Online: " . onlineCount, 1)
    SB_SetText("Offline: " . offlineCount, 2)
    SB_SetText("Next scan in " . scantimer, 3)
    return

;<=====  Functions  ===========================================================>
CheckHosts(){
    Global
    SetTimer, UpdateSB, Off
    SetTimer, CheckHosts, Off
    SB_SetText("Online: 0", 1)
    SB_SetText("Offline: 0", 2)
    SB_SetText("Scan in progress...", 3)
    onlineCount := 0
    offlineCount := 0
    threads := Object()
    Loop, % hosts.MaxIndex()
    {
        ;Wait for free thread
        while (threads.MaxIndex() >= settings.selectSingleNode("/hostMonitor/settings/maxThreads").text)
            sleep, 250

        ;Start thread to check host
        if A_IsCompiled {
            Run, % A_ScriptDir . "\inc\PingMsg.exe """ . A_ScriptName . """ "
                . hosts[A_Index, "name"] . " " . A_Index,, Hide, threadID
        }
        else {
            Run, % A_ScriptDir . "\inc\PingMsg.ahk """ . A_ScriptName . """ "
                . hosts[A_Index, "name"] . " " . A_Index,, Hide, threadID
        }

        hosts[A_Index, "threadID"] := threadID
        threads.Push(threadID)

        GuiControl,, % hosts[A_Index, "statusTextID"], % "Checking host ("
            . hosts[A_Index, "threadID"] . ")"

        ;Update status bar
        SB_SetText(threads.MaxIndex() . " threads active", 3)
    }
    scanTimer := settings.selectSingleNode("/hostMonitor/settings/checkInterval").text
    SetTimer, CheckHosts, % (settings.selectSingleNode("/hostMonitor/settings/checkInterval").text * 1000)
    SetTimer, UpdateSB, On
    SetTimer, StaleThreads, % (((settings.selectSingleNode("/hostMonitor/settings/checkInterval").text) / 2) * 1000)
}

;FlushDNS by jNizM
FlushDNS(){
    if !(DllCall("dnsapi.dll\DnsFlushResolverCache"))
    {
        throw Exception("DnsFlushResolverCache", -1)
        return 0
    }
    return 1
}

LoadXML(file){
    xmlFile := fileOpen(file, "r")
    xml := xmlFile.read()
    xmlFile.Close()

    doc := ComObjCreate("MSXML2.DOMdocument.6.0")
    doc.async := false
    if !doc.loadXML(xml)
    {
        MsgBox, % "Could not load" . file . "!"
    }

    return doc
}

Log(fileName, text){
    FormatTime, TimeStamp, A_Now, [dd/MMM/yyyy HH:mm:ss]
    logFile := fileOpen(fileName, "a")
    logFile.Write(TimeStamp . " " . text . "`r`n")
    logFile.Close()
    return 1
}

LogStart(settings){
    FormatTime, TimeStamp, A_Now, dd-MMM-yyyy_HH-mm-ss
    IfNotExist, % A_ScriptDir . "\Logs\"
        FileCreateDir, % A_ScriptDir . "\Logs"
    fileName := A_ScriptDir . "\Logs\" . TimeStamp . ".txt"
    Try {
        logFile := fileOpen(fileName, "w")
    }
    catch e {
        MsgBox, Failed to open file for logging!`n%A_LastError%
    }
    logFile.Write("[" . TimeStamp . "] Logging started. Using hosts file " . settings.selectSingleNode("/hostMonitor/settings/hostPath").text . "`r`n")
    logFile.Close()
    return fileName
}

LogStop(fileName){
    FormatTime, TimeStamp, A_Now, [dd/MMM/yyyy HH:mm:ss]
    logFile := fileOpen(fileName, "a")
    logFile.Write(TimeStamp . " Logging finished.`r`n")
    logFile.Close()
    return 1
}

Receive_WM_COPYDATA(wParam, lParam){
    Global
    Critical
    StringAddress := NumGet(lParam + 2*A_PtrSize)
    CopyOfData := StrGet(StringAddress)

    ;hostID|HostName|PingTime|IP
    reply := StrSplit(CopyOfData, "|")

    ;Clear thread from threads array
    Loop, % threads.MaxIndex()
    {
        if (threads[A_Index] == hosts[reply[1], "threadID"])
            threads.removeAt(A_Index)
    }

    ;Process reply
    if (reply[3] != "TIMEOUT")
    {
        ;Good return
        ;Check if host's ping array is full, remove oldest element if so
        if (hosts[reply[1]].pingArray.getCapacity() == settings.selectSingleNode("/hostMonitor/settings/pingAvgCount").text)
            hosts[reply[1]].pingArray.RemoveAt(1)

        ;Add new ping time to array
        hosts[reply[1]].pingArray.Push(reply[3])

        ;Update host's name element if its an IP & update GUI
        if (RegExMatch(hosts[reply[1], "name"], "^((|\.)\d{1,3}){4}$") && !hosts[reply[1], "alias"])
        {
            hosts[reply[1], "name"] := reply[2]
            GuiControl,, % hosts[reply[1], "nameTextID"], % subStr(hosts[reply[1], "name"], 1, 25)
        }

        ;Update host's ip element if different
        if (reply[4] != hosts[reply[1], "ip"])
            hosts[reply[1], "ip"] := reply[4]

        ;Update host's lastSeen element
        hosts[reply[1], "lastSeen"] := A_Hour . ":" . A_Min

        ;Calculate average ping time from array
        sum := 0
        for each, ping in hosts[reply[1]].pingArray
            sum += ping
        avg := floor(sum/hosts[reply[1]].pingArray.getCapacity())
        SetFormat, FloatFast, 0.0

        ;Update BGImage. Yellow if above warnLatency, green otherwise
        if (strReplace(reply[3], "ms") >= settings.selectSingleNode("/hostMonitor/settings/warnLatency").text)
        {
            ;GuiControl,, % hosts[reply[1], "bgImageID"], % "HBITMAP:*" hYellow
            GuiControl, +cYellow, % hosts[reply[1], "bgImageID"],

            ;Log warnings to file if enabled.
            if settings.selectSingleNode("/hostMonitor/settings/logging").text
                Log(logFile, hosts[reply[1], "name"] . " latency high: " . reply[3] . "ms (" settings.selectSingleNode("/hostMonitor/settings/warnLatency").text . "ms threshold)")
        }
        else
        {
            ;GuiControl,, % hosts[reply[1], "bgImageID"], % "HBITMAP:*" hGreen
            GuiControl, +cLime, % hosts[reply[1], "bgImageID"],
        }

        ;Refresh host text
        if hosts[reply[1], "alias"] {
            GuiControl,, % hosts[reply[1], "nameTextID"], % hosts[reply[1], "alias"]
        } else {
            GuiControl,, % hosts[reply[1], "nameTextID"], % hosts[reply[1], "name"]
        }

        ;Update status text
        GuiControl,, % hosts[reply[1], "statusTextID"], % reply[3] . "ms (" . avg . "ms)"

        ;Increment onlineCount variable
        onlineCount++
    } else {
        ;Timeout
        ;Check if host's ping array is full, remove oldest element if so
        if (hosts[reply[1]].pingArray.getCapacity() == settings.selectSingleNode("/hostMonitor/settings/pingAvgCount").text)
            hosts[reply[1]].pingArray.RemoveAt(1)

        ;Add new ping time to array
        hosts[reply[1]].pingArray.Push(9999)

        ;Update BGImage to red
        ;GuiControl,, % hosts[reply[1], "bgImageID"], % "HBITMAP:*" hRed
        GuiControl, +cRed, % hosts[reply[1], "bgImageID"],

        ;Refresh host text
        if hosts[reply[1], "alias"] {
            GuiControl,, % hosts[reply[1], "nameTextID"], % hosts[reply[1], "alias"]
        } else {
            GuiControl,, % hosts[reply[1], "nameTextID"], % hosts[reply[1], "name"]
        }

        ;Update status text
        GuiControl,, % hosts[reply[1], "statusTextID"], % "Last Seen: " . hosts[reply[1], "lastSeen"]

        ;Log failures to file if enabled
        if settings.selectSingleNode("/hostMonitor/settings/logging").text
            Log(logFile, hosts[reply[1], "name"] . " TIMEOUT")

        ;Increment offlineCount variable
        offlineCount++

        ;Start tracert if autoTraceRt is enabled in settings
        if settings.selectSingleNode("/hostMonitor/settings/autoTraceRt").text
           StartTrace(hosts[reply[1], "name"])
    }

    ;Update status bar
    if !threads.MaxIndex()
        SetTimer, UpdateSB, On
    else
        SB_SetText(threads.MaxIndex() . " threads active", 3)

    ;Done...
    return true
}

SaveSettings(settings, tdoc){
    try {
        resultDoc := ComObjCreate("MSXML2.DOMdocument.6.0")
        settings.transformNodeToObject(tdoc, resultDoc)
        resultDoc.save(A_ScriptDir . "\inc\settings.xml")
        ObjRelease(Object(resultDoc))
    } catch e {
        MsgBox, % "Could not save settings.`n" . e . "`n" . e.description
        return 0
    }
    return 1
}

StaleThreads(){
    Global
    ;Loop hosts array
    for host in hosts
    {
        ;Store index in hostID for use in nested loop
        hostID := A_Index

        ;Get host thread ID
        hostThreadID := hosts[hostID, "threadID"]

        ;Check if hosts threadID is still in the threads array
        for each, tid in threads
        {
            ;if found, color the BG yellow & set text
            if (hostThreadID == tid)
            {
                ;GuiControl,, % hosts[hostID, "bgImageID"], % "HBITMAP:*" hYellow
                GuiControl, +cYellow, % hosts[hostID, "bgImageID"],

               ;Refresh host text
                if hosts[hostID, "alias"] {
                    GuiControl,, % hosts[hostID, "nameTextID"], % hosts[reply[1], "alias"]
                } else {
                    GuiControl,, % hosts[hostID, "nameTextID"], % hosts[reply[1], "name"]
                }

                ;Update status text
                GuiControl,, % hosts[hostID, "statusTextID"], % "Thread lost..."
                threads.RemoveAt(A_Index)
            }
        }
    }
    ;Release threads array
    ObjRelease(threads)
}

StartTrace(host){
    if A_IsCompiled
        Run, % A_ScriptDir . "\inc\tracert.exe " . host
    else
        Run, % A_ScriptDir . "\inc\tracert.ahk " . host
}

StartPingGraph(host){
    if A_IsCompiled
        Run, % A_ScriptDir . "\inc\pingGraph.exe " . host
    else
        Run, % A_ScriptDir . "\inc\pingGraph.ahk " . host
}

StartPutty(host, mode){
    Global
    IfNotExist, % A_ScriptDir . ".\inc\putty.exe"
    {
        MsgBox, 4,, % "This function requires putty.exe to be in the inc folder.`nGo to the putty download page?"
        IfMsgBox, Yes
            Run, http://www.chiark.greenend.org.uk/~sgtatham/putty/download.html
        else
            return
    }
    objShell := ComObjCreate("WScript.Shell")
    objExec := objShell.Exec(ComSpec . " /c .\inc\putty.exe -" . mode . " "
        . host . " " . ((mode == "ssh")?22:(mode == "telnet")?23:""))
}

WM_MOUSEMOVE() {
    global hosts
    static CurrControl, PrevControl
    CurrControl := A_GuiControl
    If (CurrControl <> PrevControl)
    {
        ToolTip  ; Turn off any previous tooltip.
        SetTimer, DisplayToolTip, 1000
        PrevControl := CurrControl
    }
    return

    DisplayToolTip:
        SetTimer, DisplayToolTip, Off
        ToolTip, % hosts[subStr(CurrControl, 2), "ip"]
        SetTimer, RemoveToolTip, 3500
        return

    RemoveToolTip:
        SetTimer, RemoveToolTip, Off
        ToolTip
        return
}

;<=====  Includes  ============================================================>
#Include %A_ScriptDir%\inc
#Include Common Functions.ahk
#Include IPHelper.ahk

