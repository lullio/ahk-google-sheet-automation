#SingleInstance, Force
settimer, fetch, -50
doc := "15PHCe8pYJjkJt0oQRn6b64lQYVpfziEKCNktPm-OzBI"
sht := "0"
lst := ""
InputBox, needle,search, , , 175, 105
ifequal, needle, , exitapp

row := []
gui, margin,0,0
gui,add, listview, x1 y1 w800 h100 vmylv grid, id_atendimento|id_paciente|id_profissional|nome_especialidade|modo_atendimento|objetivo_atendimento|tempo_total_chamada_minutos|sexo_paciente|uf_paciente|idade|nps_medico
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