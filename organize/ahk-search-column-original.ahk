doc := "1aBohx1LumhF6UICZgnI6iao8YjgK72_qnfU8O-_szqo"
sht := "731134197"
settimer, fetch, -50

lst := "", cnt := 0
InputBox, needle,search, , , 175, 105,,,,,                      ;ex. name, position, or college
if !needle
	exitapp
for x,y in strsplit(substr(oVar, 1, instr(oVar,"`r")-1),",")
	(needle = y) && pos := X

gui, margin,0,0
Gui, add, statusbar
gui, add, listview, x1 y1 w275 r10 grid vMyLV, % needle
GuiControl, -Redraw, MyLV
for x,y in strsplit(oVar,"`n","`r")
	loop, parse, y, CSV
		if (x>1 && a_index = pos)
			LV_add("",a_loopfield)
SB_SetText(LV_GetCount() " match(es)")
gui, show, , NBA Google Sheets
LV_ModifyCol(1,"AutoHdr")
GuiControl, +Redraw, MyLV
return


guiclose:
exitapp

fetch:
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("GET", "https://docs.google.com/spreadsheets/d/" doc "/export?format=csv&id=" doc "&gid=" sht, true)
whr.Send()
whr.WaitForResponse()
oVar := whr.ResponseText
return