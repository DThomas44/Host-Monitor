#SingleInstance Force
#NoEnv

Seed = C0mpl3xPa55w0rd!
String = My.Account.Name

Coded := Code(String,Seed)

MsgBox % "Coded: `n" Coded "`nUncoded: `n" Uncode(Coded,Seed)

ExitApp

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

ESC::ExitApp
