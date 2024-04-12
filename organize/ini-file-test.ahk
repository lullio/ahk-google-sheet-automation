#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

iniPath = config.ini

 ; MsgBox, Open Menu was clicked
 Gui, ConfigFile:Font, S11
 Gui, ConfigFile:New, +AlwaysOnTop -Resize -MinimizeBox -MaximizeBox, Alterar Configurações
 /*
     * COLUNA 1
 */
 Gui, ConfigFile:Add, Text,center h20 +0x200 section, Alterar Link da Planilha:
 Gui ConfigFile:Add, ComboBox, y+5 w200 center vPlanilhaLink hwndDimensoesID , Documentações Analytics|Documentações Programação|Cursos|Relatórios

 Gui, ConfigFile:Add, Text,center h20 +0x200, Nome/ID da aba da Planilha(Worksheet)
 Gui, ConfigFile:Add, Edit, vPlanilhaNomeId w200 y+5

 /*
     * COLUNA 2
 */
 Gui, ConfigFile:Add, Text, ys x+5 center h20 +0x200, Tipo de Exportação:
 Gui, ConfigFile:Add, ComboBox, vPlanilhaTipoExportacao w100 hwndCursosIDAll y+5 w200 center, CSV||HTML|JSON
 Gui, ConfigFile:Add, Text, center h20 +0x200, Range de Dados:
 Gui, ConfigFile:Add, Edit, vPlanilhaRange w200 y+5
 /*
     * FORA DAS COLUNAS
 */
 
 Gui, ConfigFile:Add, Text, xs y+10 center h20 +0x200, Query: 
 Gui, ConfigFile:Add, Edit, vPlanilhaQuery w420 y+5 r2, 

 gui, font, S13 ;Change font size to 12
 gui, ConfigFile:Add, Button, center y+15 w100 h25 Default gSaveToIniFile, &Salvar
 Gui, ConfigFile:Show, xCenter yCenter
 ControlFocus, Edit1, Cadastrar Nova Doc - Felipe Lullio

 ReadIniFile:
 ; Link da Planilha
 IniRead, PlanilhaLink, %iniPath%, planilha, linkPlanilha
 GuiControl, ConfigFile:Choose, PlanilhaLink, %PlanilhaLink%
 ; Tipo de Exportação
 IniRead, PlanilhaTipoExportacao, %iniPath%, planilha, tipoExportacao
 GuiControl, ConfigFile:Choose, PlanilhaTipoExportacao, %PlanilhaTipoExportacao%
 ; Aba da Planilha
 IniRead, PlanilhaNomeId, %iniPath%, planilha, abaPlanilha
 GuiControl, ConfigFile:Text, PlanilhaNomeId, %PlanilhaNomeId%
 ; Range de Dados
 IniRead, PlanilhaRange, %iniPath%, planilha, rangePlanilha
 GuiControl, ConfigFile:Text, PlanilhaRange, %PlanilhaRange%
 ; Query
 IniRead, PlanilhaQuery, %iniPath%, planilha, queryPlanilha
 GuiControl, ConfigFile:Text, PlanilhaQuery, %PlanilhaQuery%
 Return

 SaveToIniFile:
 Gui Submit, NoHide
  ; Link da Planilha
 IniWrite, %PlanilhaLink%, %iniPath%, planilha, linkPlanilha
  ; Tipo de Exportação
 IniWrite, %PlanilhaTipoExportacao%, %iniPath%, planilha, tipoExportacao
  ; Nome/ID da Aba
 IniWrite, %PlanilhaNomeId%, %iniPath%, planilha, abaPlanilha
  ; Range da Planilha
 IniWrite, %PlanilhaRange%, %iniPath%, planilha, rangePlanilha
  ; Query da Planilha
 IniWrite, %PlanilhaQuery%, %iniPath%, planilha, queryPlanilha
 Return