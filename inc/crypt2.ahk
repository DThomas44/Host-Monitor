;From https://autohotkey.com/board/topic/94631-encrypt-password-protected-powerful-text-encryption/page-2

String := "Each script is a plain text file containing commands to be executed by the program (AutoHotkey.exe). A script may also contain hotkeys and hotstrings, or even consist entirely of them. However, in the absence of hotkeys and hotstrings, a script will perform its commands sequentially from top to bottom the moment it is launched."

Key := "Creating a script"

Coded := XOR_String_Plus(String, Key)
Decoded := XOR_String_Minus(Coded, Key)

MsgBox % "String:`n" String "`n`nCoded:`n" Coded "`n`nDecoded:`n" Decoded

XOR_String_Plus(String,Key)
{
    Key_Pos := 1
    Loop, Parse, String
    {
        String_XOR .= Chr((Asc(A_LoopField) ^ Asc(SubStr(Key,Key_Pos,1))) + 15000)
        Key_Pos += 1
        if (Key_Pos > StrLen(Key))
            Key_Pos := 1
    }
    return String_XOR
}

XOR_String_Minus(String,Key)
{
    Key_Pos := 1
    Loop, Parse, String
    {
        String_XOR .= Chr(((Asc(A_LoopField) - 15000) ^ Asc(SubStr(Key,Key_Pos,1))))
        Key_Pos += 1
        if (Key_Pos > StrLen(Key))
            Key_Pos := 1
    }
    return String_XOR
}
