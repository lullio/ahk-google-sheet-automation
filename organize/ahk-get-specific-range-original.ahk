settimer, fetch, -50

doc := "1aBohx1LumhF6UICZgnI6iao8YjgK72_qnfU8O-_szqo", sht := "731134197"

InputBox, needle,enter cell address, , , 175, 105
ifequal, needle, , exitapp

col := regexreplace(needle,"(\D+)\d+","$L1")
row := regexreplace(needle,"\D+(\d+)","$1")

Loop, Parse, col
	{
    tmp := Asc(A_LoopField) - 96
    col += tmp * (26 ** (StrLen(col) - A_Index))
	}
for x,y in strsplit(oVar,"`n","`r")
	if (x = row)
		loop, parse, y, CSV
			if (a_index = col)
				res := a_loopfield
msgbox % res
return

fetch:
URLDownloadToFile,% "https://docs.google.com/spreadsheets/d/" doc "/export?format=csv&id=" doc "&gid=" sht, tmp.csv
while !FileExist(a_scriptdir "\tmp.csv")
	sleep, 50
fileread, oVar, tmp.csv																				; change to your needs
FileDelete, tmp.csv
return