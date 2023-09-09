settimer, fetch, -50
doc := "1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34"
sht := "0"
lst := ""
InputBox, needle,search, , , 175, 105
ifequal, needle, , exitapp

row := []
gui, margin,0,0
gui,add, listview, x1 y1 w800 h100 vmylv grid,ID|Nome|URL|Categoria
for x,y in strsplit(oVar,"`n","`r")
	{
		msgbox %x% ; Index of row
		msgbox %y% ; Value of row
	if instr(y,"Nome")
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
URLDownloadToFile, % "https://docs.google.com/spreadsheets/d/" doc "/export?format=csv&id=" doc "&gid=" sht, tmp.csv
while !FileExist(a_scriptdir "\tmp.csv")
	sleep, 100
fileread, oVar, tmp.csv										; change to your needs
FileDelete, tmp.csv
return