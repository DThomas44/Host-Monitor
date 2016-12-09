/*
    Shares.ahk
    Author: Grendahl

    This script takes  host name, host login name, and host password as
    parameters and provides the user with a list of shares on the host. 
	The user can then select a share and open it using the passed in credentials.
*/
;<=====  System Settings  =====================================================>
#SingleInstance Off
#NoEnv
#NoTrayIcon
SendMode Input

;<=====  Startup  ==========================================================>
;Set Encryption seed
Seed = C0mpl3xPas5word!

;Read in settings.xml and transform.xslt
settings := loadXML(A_ScriptDir . "\settings.xml")
tdoc := loadXML(A_ScriptDir . "\transform.xslt")

;Set host to passed in parameter
Host = %1%
;Ensure Host is actually passed in
If(!Host) {
	Msgbox Invalid Input - Must pass Host as command line parameter
	ExitApp
}
;********NEED TO CONVERT IP BACK TO NAME HERE! (or pass the name and not the IP from the context menu handler)********

StringUpper, Host, Host

;Read settings.xml
encryptedUser := settings.selectSingleNode("/hostMonitor/settings/userName").text
encryptedPass := settings.selectSingleNode("/hostMonitor/settings/userPass").text
useCreds := settings.selectSingleNode("/hostMonitor/settings/useCreds").text

;If there userName or userPass settings don't exist, go make some
If(!encryptedUser || !encryptedPass)
	gosub SaveCreds

;If User/Pass are not encrypted, send user to set them.
If encryptedUser is not integer
	notEncrypted = 1
If encryptedPass is not integer
	notEncrypted = 1
if(notEncrypted == 1)
	gosub SaveCreds

;If User/Pass are encrypted, decrypt them for use in this script.
If encryptedUser is integer
	User := Uncode(settings.selectSingleNode("/hostMonitor/settings/userName").text,Seed)
If encryptedPass is integer
	Pass := Uncode(settings.selectSingleNode("/hostMonitor/settings/userPass").text,Seed)

;Enumerate shares on host
Drives := ListShares(Host, User, Pass)

;<=====  GUI  ==============================================================>
Menu, Tray, Icon, % A_ScriptDir . "\..\img\Host Monitor.ico"
Gui, Add, Text,,Choose a share:
Gui, Add, ListBox, vShareSelection gListBox r%Count%, % Drives
Gui, Add, Button,  gSaveCreds, Credentials
Gui, Add, Button, xp+75 yp gExplore  +default, Explore
Gui, +OwnDialogs +ToolWindow
Gui, Show, ,%Host%
return

;<=====  Labels  ==============================================================>
ListBox:
	if A_GuiEvent <> DoubleClick
		return
Explore:
	GuiControlGet, Share,,ListBox1
	Share := " \\" . Host . "\" . Share
	Gui, Destroy
	Run, explorer %Share%
	If(useCreds=1) {
		Loop 20
		{
			Sleep 500
			IfWinExist, Windows Security
			{
				WinActivate, Windows Security
				WinWaitActive, Windows Security
				Sleep 500
				SendRaw %User%
				Send {TAB}
				SendRaw %Pass%
				Send {ENTER}
				Break
			}	
		}
	}
ExitApp

SaveCreds:
	Gui, Destroy
	InputBox, User, Password, Enter username to encrypt:
	User := Code(User,Seed)
	InputBox, Pass, Password, Enter password to encrypt:, Hide
	Pass := Code(Pass,Seed)
	node := settings.selectSingleNode("/hostMonitor/settings/userName")
        node.text := User
        SaveSettings(settings, tdoc)
	node := settings.selectSingleNode("/hostMonitor/settings/userPass")
        node.text := Pass
        SaveSettings(settings, tdoc)	
	notEncrypted = 0
	Run, % A_ScriptFullPath . " " . Host
ExitApp

GuiClose:
    ExitApp

ESC::ExitApp

;<=====  Functions  ==============================================================>
ListShares(Server, User, Pass) { ;Enumerate shares on a remote host
	Global Count = 0
	PropertyList := "Name"
	wmiLocator := ComObjCreate("WbemScripting.SWbemLocator")
	try
		objWMIService := wmiLocator.ConnectServer(Server, "root\cimv2", User, Pass)
	catch e
	{	
		MsgBox % "Unable to connect to " Server "`,`n" User " does not have permission.`n`n" e.Extra " Error Message:`n"  e.Message
		ExitApp
	}
	WQLQuery = Select * From Win32_Share
	colDiskDrive := objWMIService.ExecQuery(WQLQuery)._NewEnum
	While colDiskDrive[objDiskDrive] 
		Loop, Parse, PropertyList, `,
		{	
			Count ++
			Shares  .= A_index = 1 ? objDiskDrive[A_LoopField] . "|"  :  . A_LoopField
		}
	StringTrimRight, Shares, Shares, 1
	Return, Shares
}

Code(String, Seed) { ;Encrypt a string using Mersenne Twister
	Random,, Seed
	Loop, Parse, String
	{
		Random x, 1, 1000000
		Random y, 1, 1000000
		newString .= (Asc(A_loopfield)+x) y
	}
	return newString
}

Uncode(String, Seed) { ;Decrept a Mersenne Twister encrypted string
	Random,, Seed
	while StrLen(String)>0
	{
		Random x, 1, 1000000
		Random y, 1, 1000000
		Pos := InStr(String, y)
		oldString .= Chr(SubStr(String, 1, Pos-1)-x)
		String := SubStr(String, Pos+StrLen(y))
	}	
	return oldString
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

SaveSettings(settings, tdoc){
    try {
        resultDoc := ComObjCreate("MSXML2.DOMdocument.6.0")
        settings.transformNodeToObject(tdoc, resultDoc)
        resultDoc.save(A_ScriptDir . "\settings.xml")
        ObjRelease(Object(resultDoc))
    } catch e {
        MsgBox, % "Could not save settings.`n" . e . "`n" . e.description
        return 0
    }
    return 1
}
