/*
	Common Funtions
	Author: Daniel Thomas
*/
;<=====  enumObj  =============================================================>
;Enumerates object to string. Sub elements are indented.
enumObj(Obj, indent := 0){
	if !isObject(Obj)
		MsgBox, Not an object!`n-->%obj%<--
	srtOut := ""
	for key, val in Obj
	{
		if isObject(val)
		{
			loop, %indent%
				strOut .= A_Tab
			strOut .= key . ":`n" . enumObj(val, indent + 1)
		}
		else
		{
			loop, %indent%
				strOut .= A_Tab
			strOut .= key . ": " . val . "`n"
		}
	}
	return strOut
}

;<=====  strTrim  =============================================================>
;Trims whitespace from start and end of a string.
strTrim(string){
	string = %string%
	return string
}

;<=====  padLeft  =============================================================>
;Pads a string on the left with the designated character, or spaces if not provided.
padLeft(string, len, char := "%A_Space%"){
	while (strLen(string) < len)
		string := char . string
	return string
}

;<=====  padRight  ============================================================>
;Pads a string on the right with the designated character, or spaces if not provided.
padRight(string, len, char := "%A_Space%"){
	while (strLen(string) < len)
		string .= char
	return string
}

;<===== secToTime  ============================================================>
;Converts a value in seconds to hh:mm:ss
secToTime(sec){
	strOut := ""
	if (sec >= 3600)
	{
		h := floor(sec/3600)
		sec := sec - (h*3600)
		if (strLen(h) < 2)
			h := "0" . h
		strOut := h . ":"
	}
	else
		strOut := "00:"
	if (sec >= 60)
	{
		m := floor(sec/60)
		sec := sec - (m*60)
		if (strLen(m) < 2)
			m := "0" . m
		strOut .= m . ":"
	}
	else
		strOut .= "00:"
	if (strLen(sec) < 2)
		sec := "0" . sec
	strOut .= sec
	return strOut
}
