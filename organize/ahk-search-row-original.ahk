; post: https://www.autohotkey.com/boards/viewtopic.php?t=75196
settimer, fetch, -50
doc := "1aBohx1LumhF6UICZgnI6iao8YjgK72_qnfU8O-_szqo"
sht := "731134197"
lst := ""
InputBox, needle,search, , , 175, 105
ifequal, needle, , exitapp

row := []
gui, margin,0,0
gui,add, listview, x1 y1 w800 h100 vmylv grid,name|year_start|year_end|position|height|height (f)|height (in)|height (m)|weight|weight (kg)|LMD (kg/m)|birth_date|college
for x,y in strsplit(oVar,"`n","`r")
	{
	if instr(y,needle)
		{
		loop, parse, y, CSV
			{
			row.push(a_loopfield)
			ifequal,a_index,13,break							;prevents from reading columns that are further out 
			}
		LV_add("",row*)	
		row := []
		}
	}
LV_ModifyCol()
gui, show
return

guiclose:
exitapp

fetch:
URLDownloadToFile,% "https://docs.google.com/spreadsheets/d/" doc "/export?format=csv&id=" doc "&gid=" sht, tmp.csv
while !FileExist(a_scriptdir "\tmp.csv")
	sleep, 100
fileread, oVar, tmp.csv										; change to your needs
FileDelete, tmp.csv
return