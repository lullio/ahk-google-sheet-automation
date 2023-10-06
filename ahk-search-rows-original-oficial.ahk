doc := "1aBohx1LumhF6UICZgnI6iao8YjgK72_qnfU8O-_szqo"
sht := "731134197"
settimer, fetch, -50

lst := "", cnt := 0
InputBox, needle,search, , , 175, 105,,,,,Bob
if !needle
	exitapp

gui, margin,0,0
gui, add, statusbar
gui, add, listview, x1 y1 w1175 r10 grid vMyLV,name|year_start|year_end|position|height|height (f)|height (in)|height (m)|weight|weight (kg)|LMD (kg/m)|birth_date|college
GuiControl, -Redraw, MyLV
for x,y in strsplit(oVar,"`n","`r")
	if instr(y,needle) ; se encontrar o texto digitado no searchbox na linha
		{
		row := [], ++cnt
		loop, parse, y, CSV ; dividir a linha em células
			if (a_index <= 13)																	;or if a_index in 1,4,5
				row.push(a_loopfield)
		LV_add("",row*)
		}
SB_SetText(cnt " match(es)")
gui, show, , NBA Google Sheets
loop, % lv_getcount("col")
	LV_ModifyCol(a_index,"AutoHdr")
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
