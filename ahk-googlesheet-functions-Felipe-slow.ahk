﻿/*
DICAS:
1. Não pode usar vírgula nos campos da planilha pois vai ser entendido como uma coluna em vez de linha
2. Caso queira inserir mais de um item no campo da planilha, não use vírgula ou quebra de linha para separar, use " | "
3. Os dados CSV retornam entre aspas "", isso é bom para você transformar em arrays e usar como variável javascript

*/
#Include, <Default_Settings>

full_command_line := DllCall("GetCommandLine", "str")
if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
   {
      try
      {
        if A_IsCompiled
            Run *RunAs "%A_ScriptFullPath%" /restart
        else
         Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
      }
      ExitApp
   }

Menu, Tray, Icon, C:\Windows\system32\imageres.dll,312 ;Set custom Script icon

DLLPath=C:\Users\%A_UserName%\Documents\Github\AHK\secondary-scripts\ahk-styles\styles\USkin.dll ;Location to the USkin.dll file
StylesPath=C:\Users\%A_UserName%\Documents\Github\AHK\secondary-scripts\ahk-styles\styles ;location where you saved the .msstyles files

; melhores dark: cosmo, lakrits
; melhores light: MacLion3, Milikymac, Panther, Milk, Luminous, fanta, invoice
SkinForm(DLLPath,Apply, StylesPath "\MacLion3.msstyles") ; cosmo. msstyles

; Gosub, Gui
; SkinForm(DLLPath,"0", StylesPath . CurrentStyle)

SkinForm(DLLPath,Param1 = "Apply", SkinName = ""){
  if(Param1 = Apply){
    DllCall("LoadLibrary", str, DLLPath)
    DllCall(DLLPath . "\USkinInit", Int,0, Int,0, AStr, SkinName)
  }else if(Param1 = 0){
    DllCall(DLLPath . "\USkinExit")
  }
}

/*
   *VARIÁVEIS PARA FORMAR A URL DO GOOGLE SHEET*
   - Somente a sheetURL_key é obrigatória
     
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" sheetURL_key "gviz/tq?tqx=out:" sheetURL_format "&range=" sheetURL_range "&sheet=" sheetURL_name "&tq=" sheetURL_SQLQueryEncoded
   msgbox % fullSheetURL 
*/

sheetURL_key := "1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34" ; id da pasta de trabalho/arquivo
sheetURL_name := "All-Docs" ; nome ou id da aba / guia / planilha
sheetURL_format := "csv" ; csv, html ou json
sheetURL_range := "" ; A1:C99
sheetURL_SQLQueryGA4Doc := "select * where D matches '^GA4.*' AND D is not null"
sheetURL_SQLQuery := "select * where A matches '.*' AND A is not null"
sheetURL_SQLQueryEncoded = % GS_EncodeDecodeURI(sheetURL_SQLQuery)
global i:=1 ; contas quantas vezes clicou no botão (botão Pesquisar)
/*
* ; VARIÁVEIS INI (ARQUIVO DE CONFIGURAÇÃO)

*/
if((A_PtrSize=8&&A_IsCompiled="")||!A_IsUnicode){ ;32 bit=4  ;64 bit=8
   SplitPath,A_AhkPath,,dir
   if(!FileExist(correct:=dir "\AutoHotkeyU32.exe")){
      MsgBox error
      ExitApp
   }
   Run,"%correct%" "%A_ScriptName%",%A_ScriptDir%
   ExitApp
}

if !InStr(A_OSVersion, "10.")
  appdata := A_ScriptDir
else
  appdata := A_AppData "\" regexreplace(A_ScriptName, "\.\w+"), isWin10 := true
iniPath = %appdata%\settings.ini
Run, %iniPath%

; msgbox % capture_sheetURL_key1 ; 1 serve para retornar o 1º capturing group(o que está entre parênteses )
; msgbox % capture_sheetURL_name1 ; 1 serve para retornar o 1º capturing group(o que está entre parênteses )

/*
   * CRIAR A GUI
   * CONTROLS
   *
   *
*/

Gui, Destroy
Gui,+AlwaysOnTop ; +Owner
gui, font, S11 ;Change font size to 12
/*
MENU BAR
*/
Menu, FileMenu, Add, &Abrir Planilha`tCtrl+O, MenuAbrirLink
Menu, FileMenu, Add, &Abrir Pasta Drive`tCtrl+D, MenuAbrirLink
Menu, FileMenu, Add, &Abrir Pasta Script, MenuAbrirLink

Menu, EditMenu, Add, Trocar Planilha e/ou Configurações`tCtrl+E, MenuEditarBase
; Menu, EditMenu, Add, Trocar Planilha(Aba), MenuEditarBase
; Menu, EditMenu, Add, Alterar Formato de Exportação`tCtrl+A, MenuEditarBase
; Menu, EditMenu, Add, Alterar Range de Dados`tCtrl+A, MenuEditarBase
; Menu, EditMenu, Add, Definir query para Planilha`tCtrl+A, MenuEditarBase
Menu, EditMenu, Add ; with no more options, this is a seperator
Menu, EditMenu, Add, &Reiniciar o App`tCtrl+R, MenuAcoesApp
Menu, EditMenu, Add, &Sair do App`tCtrl+Esc, MenuAcoesApp


Menu, HelpMenu, Add, &Sobre o programa, MenuAbrirLink
Menu, HelpMenu, Add, &Desenvolvedor, MenuAbrirLink
Menu, HelpMenu, Add, &WhatsApp, MenuAbrirLink

; Attach the sub-menus that were created above.
Menu, MyMenuBar, Add, &Arquivo, :FileMenu
; Menu, MyMenuBar, Add, &Editar, :EditMenu
; Menu, MyMenuBar, Add, &Editar, :EditMenu
Menu, MyMenuBar, Add, &Editar, :EditMenu
Menu, MyMenuBar, Add, &Ajuda, :HelpMenu

Gui, Menu, MyMenuBar ; Attach MyMenuBar to the GUI

/*
   STATUS BAR
*/
Gui Font, S9
Gui Add, Statusbar, gStatusBarLinks vMyStatusBar,
/*
   EDITAR TEXTO DA STATUS BAR
*/
SB_SetParts(200, 200, 100)
; SB_SetText("Total de Linhas: ", 1)

SB_SetText("Abrir Planilha", 3)
Gui Font, S9

; DEFINIR TODAS AS TABS
Gui Add, Tab3, vTabVariable gTabLabel, All|GA4|GDS|BigQ|Pixels|GTM
Gui Font, S10

; CRIAR A PRIMEIRA TAB
Gui Tab, All 
Gui, font, S10

/*
 * ********* TAB 1
*/
; se quiser que apareça nome do grupo tirar o -Hdr
Gui, Add, ListView, r15 Grid NoSortHdr vLVAll w450 gListViewListener,
Gui, Add, Edit, h29 vVarPesquisarDados w230 y+10 section, GA4
Gui, Add, Button, vBtnPesquisar x+10 w100 h30 gPesquisarDados Default, Pesquisar
Gui, Add, Button, vBtnAtualizar x+10 w100 h30 gAtualizarPlanilha, Atualizar
Gui, Add, Checkbox, vCheckIdiomaPt Checked1 xs y+10 gListenerIdioma, abrir documentação em português
Gui, Add, Checkbox, vCheckPesquisarColuna Checked0 x+10, pesquisar por coluna
; Gui Add, Link, w120 x+10 vTotalLinhas center,

/*
 * ********* TAB 2
*/
Gui Tab, GA4
Gui, font,center S11 cBlue
; ! TIPOS DE EVENTOS GA4
; dropdown 1 - principais cursos
/*
-----------
----------- 1ª COLUNA
-----------
*/
Gui Add, Text, section y+15, Tipos de Eventos GA4
Gui, Add, ComboBox, w200 vGDocsEventos hwndIdEventos gDocs, 

; ! METRICAS GA3 VS GA4
; dropdown 1 - principais cursos
Gui Add, Text,, Comparar Métricas GA3 vs GA4
Gui, Add, ComboBox, w200 vGDocsMetricas hwndIdMetricas gDocs,

; !!!!!! MIGRAÇÃO DO GA3 PARA GA4 - COMPATIBILIDADE

; dropdown 1 - principais cursos
Gui Add, Text,, GA4 Migration
Gui, Add, ComboBox, w200 vGDocsMigration hwndIdMigration gDocs,
/*
-----------
----------- 2ª COLUNA
-----------
*/
; ! RELEASE NOTES

; dropdown 1 - principais cursos
Gui Add, Text,ys x+10, What's New - Release Notes
Gui, Add, ComboBox, w200 vGDocsReleases hwndIdRelease gDocs, 

; ! LIMITS

; dropdown 1 - principais cursos
Gui Add, Text,, Limits and Price
Gui, Add, ComboBox, w200 vGDocsLimits hwndIdLimits gDocs,

; ! ACCOUNT STRUCTURE

; dropdown 1 - principais cursos
Gui Add, Text,, Account Structure
Gui, Add, ComboBox, w200 vGDocsAccount hwndIdAccount gDocs, 
Gui, Add, Link, xs+90 y+20,<a>Root-Doc</a> | <a>What's New</a> | <a>Blog</a> | <a>Notion</a>
Gui, Add, Checkbox, Checked1 VIdioma x+15, pt-br?
; Botões
gui, font, S11
gui, Add, Button, xs+20 y+20 w200  vVarAbrirDoc1 gAbrirDoc Default, &Abrir Doc
gui, Add, Button, w150 x+20 Cancel gCancel, &Cancelar

/*
 * ********* TAB 3
*/
Gui Tab, GDS
/*
-----------
----------- 1ª COLUNA
-----------
*/
; ! DOC TEXT GDS

; dropdown 1 - principais cursos
Gui, Font, S11
Gui Add, Text,y+15 section, Text Functions / Regex
Gui, Add, ComboBox, w200 vGDSDocsRegex hwndIdRegex gDocs, 

; ! DOC CONDITIONAL GDS

; dropdown 1 - principais cursos
Gui Add, Text, , Conditional Functions
Gui, Add, ComboBox, w200 vGDSDocsConditional hwndIdConditional gDocs, 

; ! DOC Agregração GDS

; dropdown 1 - principais cursos
Gui Add, Text,, Aggregation Functions
Gui, Add, ComboBox, w200 vGDSDocsAggregation hwndIdAggregation gDocs, 

; ! DOC Dates GDS

/*
-----------
----------- 2ª COLUNA
-----------
*/
; dropdown 1 - principais cursos
Gui Add, Text, ys x+10, Date Functions
Gui, Add, ComboBox, w200 vGDSDocsDates hwndIdDates gDocs, 

; ! DOC CALCULATED FIELDS DOCUMENTATION

; dropdown 1 - principais cursos
Gui Add, Text,, Calculated Fields / Resource Infos
Gui, Add, ComboBox, w200 vGDSDocsCalculated hwndIdCalculated gDocs, 

; ! DOC OUTROS DOCS GDS

; dropdown 1 - principais cursos
Gui Add, Text,, Outras Docs - Blend/Parametros/Etc
Gui, Add, ComboBox, w200 vGDSDocsOutros hwndIdOutros gDocs, 

Gui, Add, Link, xs+90 y+20,<a>Root-Doc</a> | <a>What's New</a> | <a>Blog</a> | <a>Notion</a>
Gui, Add, Checkbox, Checked1 VIdiomaGDS x+15, pt-br?

; Botões
gui, font, S11
gui, Add, Button, xs+20 y+20 w200 vVarAbrirDoc2 gAbrirDoc Default, &Abrir Doc
gui, Add, Button, w150 x+20 Cancel gCancel, &Cancelar

/*
 * ********* TAB 4
*/
Gui Tab, BigQ
; ! DOC BIG QUERY GA4

; dropdown 1 - principais cursos
Gui, Font, S11
Gui Add, Text,y+15 section, BigQuery GA4 - Documentações
Gui, Add, ComboBox, section w200 vBQDoc1 hwndIdBQ1 gDocs, 

Gui, Add, Link,  y+15,<a>Root-Doc</a> | <a>What's New</a> | <a>Blog</a> | <a>Notion</a>
Gui, Add, Checkbox, Checked1 VIdiomaBQ x+15, pt-br?

; Botões
gui, font, S11
gui, Add, Button, xs+10 y+30 w100  vVarAbrirDoc3 gAbrirDoc Default, &Abrir Doc
gui, Add, Button, w75 x+10 Cancel gCancel, &Cancelar

/*
 * ********* TAB 5
*/
; obj properties: https://developers.facebook.com/docs/meta-pixel/reference#object-properties
; custom data parameters: https://developers.facebook.com/docs/marketing-api/conversions-api/parameters/custom-data#
; for Advantage+ catalog ads: https://www.facebook.com/business/help/606577526529702?id=1205376682832142
Gui Tab, Pixels

; dropdown 1 - principais cursos
gui, font, S11
Gui Add, Text,y+15 section, Facebook Pixel
Gui, Add, ComboBox, w200 vFbDocsPixel hwndIDFbPixel gDocs, 

; ! FACEBOOK API

; dropdown 1 - principais cursos
Gui Add, Text,, Facebook API de Conversões
Gui, Add, ComboBox, w200 vFbDocsAPI hwndIDFbAPi gDocs, 

; ! TIK TOK DOCUMENTATION

; dropdown 1 - principais cursos
Gui Add, Text,, TikTok Pixel
Gui, Add, ComboBox, w200 vTikTokDocsPixel hwndIdTikTok gDocs, 

; ! KWANKO DOCUMENTATION

; dropdown 1 - principais cursos
Gui Add, Text, x+10 ys, Kwanko Pixel
Gui, Add, ComboBox, w200 vKwankoDocs hwndIdKwanko gDocs, 

; ! FLOODLIGHT

; dropdown 1 - principais cursos
Gui Add, Text,, FloodLight Pixel
Gui, Add, ComboBox, w200 vFloodLightDocs hwndIdFloodLight gDocs, 

; ! Outros Pixels

; dropdown 1 - principais cursos
Gui Add, Text,, Outros Pixel
Gui, Add, ComboBox, w200 vOutrosPixelDocs hwndIdOutrosPixel gDocs, 

Gui, Add, Link, xs+90 y+20,<a>Root-Doc</a> | <a>What's New</a> | <a>Blog</a> | <a>Notion</a>
Gui, Add, Checkbox, Checked1 VIdiomaPixels x+15, pt-br?

; Botões
gui, font, S11
gui, Add, Button, xs+20 y+20 w200 vVarAbrirDoc4 gAbrirDoc Default, &Abrir Doc
gui, Add, Button, w150 x+20 Cancel gCancel, &Cancelar

/*
 * ********* TAB 6
*/
Gui Tab, GTM
; ! TIPOS DE EVENTOS GA4

; dropdown 1 - principais cursos
Gui Add, Text,y+15, Camada de Dados
Gui, Add, ComboBox, w300 vGDocsGTMDL hwndIdGTMDL gDocs, 

; ! METRICAS GA3 VS GA4

; dropdown 1 - principais cursos
Gui Add, Text,y+15, Comparar Métricas GA3 vs GA4
Gui, Add, ComboBox, w300 vGDocsGTMO hwndIdGTMO gDocs, 
Gui, Add, Link,  xs y+15,<a>Root-Doc</a> | <a>What's New</a> | <a>Blog</a> | <a>Notion</a>
Gui, Add, Checkbox, Checked1 VIdiomaGTM x+15, pt-br?

; Botões
gui, font, S11
gui, Add, Button, xs+10 y+30 w100  vVarAbrirDoc5 gAbrirDoc Default, &Abrir Doc
gui, Add, Button, w75 x+10 Cancel gCancel, &Cancelar

Gui, Show, AutoSize , Web Analytics Links Helper - Felipe Lullio

/*
* AO SELECIONAR UMA TAB FOCAR NO BOTÃO "ABRIR DOC" da tab correspondente 
*/
; BOTÃO PADRÃO
GuiControl, +Default, BtnPesquisar
; FOCAR NO EDIT CONTROL DE PESQUISA
ControlFocus, Edit1, Web Analytics Links

Gui, ListView, LVAll
/*
   LER ARQUIVO DE CONFIGURAÇÃO
*/
ReadIniFile:
Gui Submit, NoHide
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
; msgbox %PlanilhaTipoExportacao%
; GS_GetCSV_ToListView()
GS_GetCSV_ToListView(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
Return
/*
   ESCREVER NO ARQUIVO DE CONFIGURAÇÃO
*/
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
Notify().AddWindow("Configuração atualizada!",{Time:3000,Icon:177,Background:"0x039018",Title:"SUCESSO",TitleColor:"0xF0F8F1", TitleSize:15, Size:15, Color: "0xF0F8F1"},"","setPosBR")
Run %iniPath%
Return


/*
   VARIÁVEIS QUE CONTÉM OS VALORES DAS COLUNAS DA PRIMEIRA LINHA
*/
global ColumnCategory := GS_GetCSV_Column(, "i)Categoria",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).arrColumnSanitize ; ColumnData.variavelJavascript ColumnData.arrColumn ColumnData.strColumn
global UniqueColumnCategory := RmvDuplic(ColumnCategory)
; Msgbox % ColumnCategory.arrColumnSanitize[5]
/*
TRATAMENTO DO MENU BAR
*/
MenuHandler:
; MsgBox, %A_ThisMenuItem%
return

Docs:
Return
/* TRATAMENTO DOS DROPDOWN, PARA QUANDO VC ESCREVER O NOME DO CURSO JÁ PREENCHER O CURSO AUTOMATICAMENTE NO DROPDOWN
*/
; RESOLVI CRIAR UMA FUNÇÃO PARA NÃO TER QUE DUPLICAR ESSE CÓDIGO VÁRIAS VEZES PARA OS DROPDOWNS
DropDownComplete(DocID)
{
  ControlGetText, Eingabe,, ahk_id %DocID%
  ControlGet, Liste, List, , , ahk_id %DocID%
  ; msgbox %Liste%
  ; msgbox %Eingabe%
  ; If ( !GetKeyState("Delete") && !GetKeyState("BackSpace") && RegExMatch(Liste, "`nmi)^(www\.)?(\Q" . Eingabe . "\E.*)$", Match)) {
  If ( !GetKeyState("Delete") && !GetKeyState("BackSpace") && RegExMatch(Liste, "`nmi)^(\Q" . Eingabe . "\E.*)$", Match)) {
    ; msgbox %match%
    ; msgbox %match1% ; armazena o www.
    ; msgbox %match2% ; armazena o restante sem o www.
    ControlSetText, , %Match%, ahk_id %DocID% ; insere o texto no combobox
    Selection := StrLen(Eingabe) | 0xFFFF0000 ; tamanho do texto do match
    ; msgbox %Selection%
    SendMessage, CB_SETEDITSEL := 0x142, , Selection, , ahk_id %DocID% ; colocar o Docr do mouse selecionando o texto do match
  } Else {
    CheckDelKey = 0
    CheckBackspaceKey = 0
  }
  ; GuiControl,Focus,Curso
}


; GoSub, controlVideos
; Ignorar o erro que o ahk dá e continuar executando o script





/*
   *FUNÇÃO PARA DECODIFICAR A QUERY QUE VAI NA URL*
   ; https://autohotkey.com/board/topic/17367-url-encoding-and-decoding-of-special-characters/
   ; https://developers.google.com/chart/interactive/docs/querylanguage?hl=pt-br#language-clauses

   # Exemplo de uso
   sheetURL_SQLQuery := "select A, sum(B) group by A"
   MsgBox, % decoded := GS_EncodeDecodeURI(sheetURL_SQLQuery, false)
   MsgBox, % GS_EncodeDecodeURI(decoded)
*/
GS_EncodeDecodeURI(str, encode := true, component := true) {
   static Doc, JS
   if !Doc {
      Doc := ComObjCreate("htmlfile")
      Doc.write("<meta http-equiv=""X-UA-Compatible"" content=""IE=9"">")
      JS := Doc.parentWindow
      ( Doc.documentMode < 9 && JS.execScript() )
   }
   Return JS[ (encode ? "en" : "de") . "codeURI" . (component ? "Component" : "") ](str)
}

/*
   * FUNÇÃO PARA RETORNAR OS DADOS DA PLANILHA, RETORNAR A TABELA
   - Somente a sheetURL_key é obrigatória mas eu já deixei um valor padrão nela que é a planilha "Automate Documentations"
   # Para testar:
   msgbox % GS_GetCSV()

*/
GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   Gui Submit, NoHide
   /*
      * capturar o nome da planilha pela gui/arquivo de configuração .ini
      * se o valor "abaPlanilha" estiver vazio no arquivo de configuração, capturar o nome da planilha pela URL da planilha.
   */
   RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   If(PlanilhaNomeId)
      capture_sheetURL_name := PlanilhaNomeId
   Else
      capture_sheetURL_name := capture_sheetURL_name1
   ; msgbox % capture_sheetURL_name
   ; msgbox % capture_sheetURL_key1
   ; msgbox % capture_sheetURL_name1
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key1 "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&sheet=" capture_sheetURL_name "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)

   whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
   whr.Open("GET",fullSheetURL, true)
   whr.Send()
   ; Using 'true' above and the call below allows the script to remain responsive.
   whr.WaitForResponse()
   googleSheetData := whr.ResponseText
   SemAspa := RegExReplace(googleSheetData, aspa , "")
   ; Return SubStr(googleSheetData, 2,-1) ; remove o primeiro e último catactere (as aspas)
   Return googleSheetData
}
/*
   * FUNÇÃO PARA CAPTURAR DADOS DE UMA COLUNA ESPECÍFICA / PESQUISAR COLUNA
*/
GS_GetCSV_Column(JS_VariableName:="arr", regexFindColumn := "i).*", PlanilhaLink:="", PlanilhaQuery:="", PlanilhaTipoExportacao:="csv", PlanilhaRange:="", PlanilhaNomeId:=""){
    Gui Submit, NoHide
    sheetData_All := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId) ; Select * limit 1
    sheetData_ColumnDataArr := []
    sheetData_ColumnDataArrSanitize := []
    sheetData_ColumnDataStr := ""
    sheetData_ColumnDataStrSanitize := ""
    sheetData_ColumnPosition := 0
    sheetData_ColumnName := ""
    sheetData_ColumnPosition := ""
   ;  regexFindColumn := "i)Categoria"

    Loop, parse, sheetData_All, `n ; PROCESSAR CADA LINHA DA TABELA/PLANILHA
       {
          LineNumber := A_Index ; Index da linha
          LineContent := A_LoopField ; Conteúdo da linha, todos valores da linha, a 1ª linha vai ser o HEADER(vc consegue capturar os headers das colunas)
       Loop, parse, A_LoopField, `, ; PROCESSAR CADA CÉLULA/CAMPO DA LINHA ATUAL
       {
         ColumnNumber := A_Index ; Index da coluna
         cellContent := A_LoopField ; armazenar o conteúdo da célula numa variável
         ; msgbox %A_LoopField% ; Exibe cada célula, cada camnpo da planilha
         ; msgbox % SubStr(A_LoopField, 2,-1) ; remove o primeiro e último catactere (as aspas)
         /*
            * Se for a linha 1 e se tiver o termo do regex na linha capture os dados da coluna somente
         */
         if(LineNumber = 1 && RegExMatch(cellContent, regexFindColumn)) ; se for a 1ª linha header e texto for igual a "nome"
         {
            sheetData_ColumnName := SubStr(cellContent, 2, -1)
            Loop, parse, sheetData_All, `n
               {
               /*
                  SALVAR TODAS AS LINHAS DA COLUNA "Nome"
               */
               ; msgbox %A_LoopField% ; aqui exibe a linha inteira (inutil)
               ; msgbox % StrSplit(A_LoopField,",")[ColumnNumber] ; exibe somente o valor da célula da coluna
               sheetData_ColumnDataArr.push(StrSplit(A_LoopField,",")[ColumnNumber])
               sheetData_ColumnDataArrSanitize.push(SubStr(StrSplit(A_LoopField,",")[ColumnNumber], 2, -1))
               sheetData_ColumnPosition := ColumnNumber
               sheetData_ColumnDataStr.= StrSplit(A_LoopField,",")[ColumnNumber] ", "
               sheetData_ColumnDataStrSanitize.= SubStr(StrSplit(A_LoopField,",")[ColumnNumber] ", ", 2, -1)
               }
            ; msgbox "Dado da coluna: " %A_LoopField%
         }
       } ; FIM DO LOOP DA COLUNA
      } ; FIM DO LOOP DA LINHA
       /*
       VARIÁVEL QUE FINALIZA A CONVERSÃO PARA UMA VARIÁVEL JAVASCRIPT
       - troca a última vírgula por ]; para finalizar a variável do tipo array
      */
       sheetData_ColumnDataStrJS = % "let " JS_VariableName " = [" RegExReplace(sheetData_ColumnDataStr, ",\s+$", "];")
       Return {variavelJavascript: sheetData_ColumnDataStrJS, arrColumn: sheetData_ColumnDataArr, arrColumnSanitize: sheetData_ColumnDataArrSanitize, strColumn: sheetData_ColumnDataStr, strColumnSanitize: sheetData_ColumnDataStrSanitize, ColumnPosition: sheetData_ColumnPosition, ColumnName: sheetData_ColumnName}
}



/*
   * FUNÇÃO PARA EXIBIR OS DADOS NA LISTVIEW
*/
GS_GetCSV_ToListView(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   Gui Submit, NoHide
   RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   ; msgbox % capture_sheetURL_key1
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   ; msgbox % capture_sheetURL_name1
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&sheet=" capture_sheetURL_name "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)
   ; msgbox %PlanilhaTipoExportacao% %PlanilhaLink% %PlanilhaNomeId% %PlanilhaRange% %PlanilhaQuery%
   sheetData_All := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   msgbox % sheetData_All
      
   ;  sheetData_All := GS_GetCSV() ; Select * limit 1

   ;  For key, index in UniqueColumnCategory
   ;    msgbox index

    Loop, parse, sheetData_All, `n ; PROCESSAR CADA LINHA DA TABELA/PLANILHA
       {
          LineNumber := A_Index ; Index da linha
          LineContent := A_LoopField ; Conteúdo da linha, todos valores da linha, a 1ª linha vai ser o HEADER(vc consegue capturar os headers das colunas)
       Loop, parse, A_LoopField, `, ; PROCESSAR CADA CÉLULA/CAMPO DA LINHA ATUAL
       {
         ColumnNumber := A_Index ; Index da coluna
         cellContent := A_LoopField ; armazenar o conteúdo da célula numa variável
         ; msgbox %A_LoopField% ; Exibe cada célula, cada camnpo da planilha
         ; msgbox % SubStr(A_LoopField, 2,-1) ; remove o primeiro e último catactere (as aspas)
       } ; FIM DO LOOP DA COLUNA
         totalColunas := ColumnNumber
         /*
           *AUTOMATIZAR A INSERÇÃO DAS LINHAS E COLUNAS
         */
         sheetData_ColumnHeaderStr := ""
         aspa := """"
         Loop, %totalColunas%
         {
            Coluna%A_Index% := RegExReplace(StrSplit(A_LoopField,",")[A_Index], aspa , "")
            ; sheetData_ColumnHeaderStr .= Coluna%A_Index% ; versão com aspas
            sheetData_ColumnHeaderStr .= Coluna%A_Index% ; versão sem aspas
            if(A_Index != totalColunas) ; se for o último índice não adicionar vírgula, para não ficar uma vírgula sozinha no final
               sheetData_ColumnHeaderStr .= ","
            ; inserir as colunas
            If(LineNumber = 1) ; adicionar as colunas com base na primeira linha
            {
              LV_InsertCol(A_Index, "center auto", Coluna%A_Index%)
            ;   msgbox %A_LoopField%
              ColunaHeader%A_Index% := SubStr(StrSplit(A_LoopField,",")[A_Index], 2, -1)
            }
         }
         If(LineNumber != 1) ; adicionar todas as linhas menos a primeira
            LV_Add("" , Coluna1, Coluna2, Coluna3, Coluna4, Coluna5, Coluna6, Coluna7, Coluna8, Coluna9, Coluna10, Coluna11, Coluna12, Coluna13, Coluna14, Coluna15, Coluna16, Coluna17, Coluna18, Coluna19, Coluna20) 
         ; msgbox %sheetData_ColumnHeaderStr%
         ;  Coluna1 := RegExReplace(StrSplit(A_LoopField,",")[1], aspa , "") ; 1ª coluna da planilha
         ; LV_Add("" , Coluna1, Coluna2, Coluna3, Coluna4) ; manter as aspas
         ; LV_Add("" , SubStr(Coluna1, 2,-1), SubStr(Coluna2, 2,-1), SubStr(Coluna3, 2,-1), SubStr(Coluna4, 2,-1), SubStr(Coluna5, 2,-1)) ; remover as aspas        
         /*
            O CONTEÚDO NA PLANILHA POSSUI OS TEXTOS "%idiomapt%", vamos tratar isso para não ser considerado um erro na url
         */
          For Index, NomeDocumentacao in StrSplit(Coluna3, " | ")
          {
               ;  msgbox % index " is " NomeDocumentacao 
                URLDocTratada := RegExReplace(NomeDocumentacao, "%idiomapt%", idioma)
               ;  msgbox % URLDocTratada
            ;  if(NomeDocumentacao != "URL")
            ;     Run % URLDocTratada
          }          
       } ; FIM DO LOOP DA LINHA

       LV_ModifyCol(1) 
       LV_ModifyCol(2, 250) 
       LV_ModifyCol(3) 
       
       ; total de linhas
       TotalLinhas:
         totalLines := LV_GetCount()
         GuiControl, , TotalLinhas, Total de Linhas: %totalLines%
         SB_SetText("Total de Linhas na Planilha: " totalLines, 1)
       Return {nomesColunas: coco, colunasHeader: [ColunaHeader1, ColunaHeader2, ColunaHeader3, ColunaHeader4, ColunaHeader5, ColunaHeader6, ColunaHeader7, ColunaHeader8, ColunaHeader9, ColunaHeader10, ColunaHeader11, ColunaHeader12, ColunaHeader13]}
}
; GS_GetCSV_ToListView()

/*
   * FUNÇÃO PARA CAPTURAR AÇÃO AO CLICAR NA LISTVIEW
*/
GS_GetListView_Click(idioma, regexFindColumnName:= ".*Nome.*", regexFindColumnURL := "i).*URL|Link.*", action := "openLink" ){
   Gui Submit, NoHide
   ; * CAPTURAR A LINHA SELECIONADA NA LISTVIEW
   NumeroLinhaSelecionada := LV_GetNext() 
   ; msgbox % NumeroLinhaSelecionada
   ; * Pesquisar por coluna específica
   getColumnName := GS_GetCSV_Column(, regexFindColumnName,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   getColumnURL := GS_GetCSV_Column(, regexFindColumnURL, PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)

   posicaoColunaNome := getColumnName.ColumnPosition
   posicaoColunaURL := getColumnURL.ColumnPosition
   valueColunaNome := getColumnName.ColumnName
   valueColunaURL := getColumnURL.ColumnName
   ; * CAPTURAR VALOR DA COLUNA "NOME"
   LV_GetText(TextoLVNome, NumeroLinhaSelecionada, posicaoColunaNome) 
   ; * CAPTURAR VALOR DA COLUNA "URL"
   LV_GetText(TextoLVURL, NumeroLinhaSelecionada, posicaoColunaURL) 
   ; msgbox %TextoLVNome% %TextoLVURL%

   ; msgbox % A_GuiEvent
   if(A_GuiEvent == "DoubleClick" && action = "openLink"){ ; abrir link normal
      /*
         * ABRIR OS LINKS/URLS/DOCUMENTAÇÕES NO NAVEGADOR
         ! IMPORTANTE: Caso tenha mais de um link na coluna, transformar em um array e fazer um loop para abrir os links
      */
      URLSanitized := StrReplace(TextoLVURL, "%idiomapt%", idioma)
      For Index, URL in StrSplit(URLSanitized, " | ")
         {
              Run, %URL%
         } 
   }else if(A_GuiEvent == "DoubleClick" && action = "openAHKChrome"){ ; abrir ahk chrome
      ; URL := RegExReplace(TextoLVURL, "%idiomapt%", "")
      ; if !(PageInst := Chrome.GetPageByURL(URL, "contains"))
      ;    {
      ;       ChromeInst := new Chrome(profileName,URL,"--remote-debugging-port=9222 --remote-allow-origins=* --profile-directory=""Default""",chPath)
      ;       Notify().AddWindow("Não encontrei o site aberto no Chrome, Vou abrir pra você agora!",{Time:6000,Icon:28,Background:"0x900C3F",Title:"OPS!",TitleSize:15, Size:15, Color: "0xCDA089", TitleColor: "0xE1B9A4"},,"setPosBR")
      ;       Sleep, 500
      ;       contador1 := 0
      ;       while !(PageInst)
      ;       {
      ;          Sleep, 500
      ;          Notify().AddWindow("procurando instância do chrome...!",{Time:6000,Icon:28,Background:"0x1100AA",Title:"ERRO!",TitleSize:15, Size:15, Color: "0xCDA089", TitleColor: "0xE1B9A4"},,"setPosBR")
      ;          PageInst := Chrome.GetPageByURL(URL, "contains")
      ;          contador1++
      ;          if(contador1 >= 30){
      ;             PageInst.Disconnect()
      ;             break
      ;          }
      ;       }
      ;    }
      ;    Sleep, 500
      ;    ; aqui está o fix pra esperar a página carregar
      ;    PageInst := Chrome.GetPageByURL(URL, "contains")
      ;    Sleep, 500
      ; ; SUPER IMPORTANTE, ATIVAR A TAB/PÁGINA, ACTIVATE, FOCUS
      ;    PageInst.Call("Page.bringToFront")
   }else if(A_GuiEvent == "RightClick"){ ; CLIQUE COM BOTÃO DIREITO DO MOUSE
      /*
         ABRIR NOTION
      */
      if(A_UserName == "Felipe" || A_UserName == "estudos" || A_UserName == "Estudos")
         {
           user := A_UserName
           pass := "xrlo1010"
         }
       Else
         {
           user := "felipe.lullio@hotmail.com"
           pass := "XrLO1000@1010"
         }
      RunAs, %user%, %pass%
      ; Run, C:\Users\felipe\AppData\Local\Programs\Notion\Notion.exe 
      Run %ComSpec% /c C:\Users\felipe\AppData\Local\Programs\Notion\Notion.exe "%TextoLinhaSelecionadaNotion%", , Hide
      RunAs
      WinActivate, Notion
   }
}

/*
   * FUNÇÃO PARA CRIAR AS CATEGORIAS
*/
test(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   global ColumnCategory := GS_GetCSV_Column(, "i)Categoria",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).arrColumnSanitize ; ColumnData.variavelJavascript ColumnData.arrColumn ColumnData.strColumn
   global UniqueColumnCategory := RmvDuplic(ColumnCategory)

   sheetData_All := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId) ; Select * limit 1
   sheetData_ColumnDataArr := []
   sheetData_ColumnDataArrSanitize := []
   sheetData_ColumnDataStr := ""
   sheetData_ColumnDataStrSanitize := ""
   sheetData_ColumnPosition := 0
   sheetData_ColumnName := ""
   sheetData_ColumnPosition := ""
   ;  regexFindColumn := "i)Categoria"

   for key, category in UniqueColumnCategory
   {
      category%key%%category%Names := []
      Loop, parse, sheetData_All, `n ; PROCESSAR CADA LINHA DA TABELA/PLANILHA
      {
         LineNumber := A_Index ; Index da linha
         LineContent := A_LoopField ; Conteúdo da linha, todos valores da linha, a 1ª linha vai ser o HEADER(vc consegue capturar os headers das colunas)
      Loop, parse, A_LoopField, `, ; PROCESSAR CADA CÉLULA/CAMPO DA LINHA ATUAL
      {
      ColumnNumber := A_Index ; Index da coluna
      cellContent := A_LoopField ; armazenar o conteúdo da célula numa variável
      ; msgbox %A_LoopField% ; Exibe cada célula, cada camnpo da planilha
      ; msgbox % SubStr(A_LoopField, 2,-1) ; remove o primeiro e último catactere (as aspas)
      /*
         * Se for a linha 1 e se tiver o termo do regex na linha capture os dados da coluna somente
      */
      msgbox % cellContent
      if(RegExMatch(category, cellContent)) ; se for a 1ª linha header e texto for igual a "nome"
      {
            ; msgbox %cellContent%
         sheetData_ColumnName := SubStr(cellContent, 2, -1)
         category%key%%category%Names.push() 
         Loop, parse, sheetData_All, `n
            {
            /*
               SALVAR TODAS AS LINHAS DA COLUNA "Nome"
            */
            ; msgbox %A_LoopField% ; aqui exibe a linha inteira (inutil)
            ; msgbox % StrSplit(A_LoopField,",")[ColumnNumber] ; exibe somente o valor da célula da coluna
            sheetData_ColumnDataArr.push(StrSplit(A_LoopField,",")[ColumnNumber])
            sheetData_ColumnDataArrSanitize.push(SubStr(StrSplit(A_LoopField,",")[ColumnNumber], 2, -1))
            sheetData_ColumnPosition := ColumnNumber
            sheetData_ColumnDataStr.= StrSplit(A_LoopField,",")[ColumnNumber] ", "
            sheetData_ColumnDataStrSanitize.= SubStr(StrSplit(A_LoopField,",")[ColumnNumber] ", ", 2, -1)
            }
         ; msgbox "Dado da coluna: " %A_LoopField%
      }
      } ; FIM DO LOOP DA COLUNA
   } ; FIM DO LOOP DA LINHA
  }
}

/*
   * FUNÇÃO PARA CAPTURAR OS COMBOBOX, CAPTURAR TODOS CONTROLS AHK
*/
AHK_GetControls(searchControls := "ComboBox"){
   Gui, Submit, NoHide
   ; PEGAR TEXTOS DA PRIMEIRA E SEGUNDA COLUNA DA LISTVIEW
   WinGet, ActiveControlList, ControlList, A
   Loop, % LV_GetCount() ; loop through every row
   {
      LV_GetText(TextoColuna1, A_Index) ; will get first column by default (Nome do Curso)
      LV_GetText(TextoColuna2, A_Index, 2) ; will get second column (URL do Curso)
      ; * CAPTURAR A LINHA SELECIONADA NA LISTVIEW
      NumeroLinhaSelecionada := LV_GetNext() 
      ; * Pesquisar por coluna específica
      getColumnName := GS_GetCSV_Column(, regexFindColumnName,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
      getColumnURL := GS_GetCSV_Column(, regexFindColumnURL,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)

      posicaoColunaNome := getColumnName.ColumnPosition
      posicaoColunaURL := getColumnURL.ColumnPosition
      valueColunaNome := getColumnName.ColumnName
      valueColunaURL := getColumnURL.ColumnName
      ; * CAPTURAR VALOR DA COLUNA "NOME"
      LV_GetText(TextoLVNome, NumeroLinhaSelecionada, posicaoColunaNome) 
      ; * CAPTURAR VALOR DA COLUNA "URL"
      LV_GetText(TextoLVURL, NumeroLinhaSelecionada, posicaoColunaURL) 
      ; msgbox %TextoLVNome% %TextoLVURL%
      /*
      CAPTURANDO TODOS OS CONTROLS DA GUI
      */
      Loop, Parse, ActiveControlList, `n
         {
         
         ControlGetText, TextoDoControl, %A_LoopField%
         FileAppend, %a_index%`t%A_LoopField%`t%TextoDoControl%`n, C:\Controls.txt
            /*
               CAPTURANDO SOMENTES OS ComboBoXES
            */
            if(InStr(A_LoopField, searchControls)) ; se for um combobox
            {
               if(TextoDoControl == "GTM1"){
                  gtm1Folder := "Y:\Season\Analyticsmania\Google Tag Manager Masterclass For Beginners 3.0"
                  if !FileExist(gtm1Folder)
                  {
                  gtm1Folder := "C:\Users\" A_UserName "\Documents\Season\Analyticsmania\Google Tag Manager Masterclass For Beginners 3.0"
                  }
                  Run vlc.exe "%gtm1Folder%\PLAYLIST-ADITIONAL-CONTENT.xspf"
                  Run %gtm1Folder%\PLAYLIST-COMPLETA-BEGGINER.xspf               
               }else if(TextoDoControl && TextoDoControl = TextoLVNome){
                  Run, %TextoLVURL%
               }
            }
         }
   }
}

/*
   * FUNÇÃO PARA PESQUISAR E RETORNAR TODAS LINHAS E COLUNAS
*/
GS_SearchRows(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   cnt := 0
   Gui Submit, NoHide
   planilha := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   GuiControl, -Redraw, LVAll
   LV_Delete()
   for x,y in strsplit(planilha,"`n","`r")
      ; if instr(y,VarPesquisarDados) ; se encontrar o texto digitado no searchbox na linha
      if RegExMatch(y, "i).*" VarPesquisarDados ".*") ; se encontrar o texto digitado no searchbox na linha
         {
         row := [], ++cnt
         loop, parse, y, CSV ; dividir a linha em células
            if (a_index <= 13)																	;or if a_index in 1,4,5
               row.push(a_loopfield)
         LV_add("",row*)
         }
   SB_SetText("Match(es) da Pesquisa: " cnt,  2)
   loop, % lv_getcount("col")
      LV_ModifyCol(a_index,"AutoHdr")
   GuiControl, +Redraw, LVAll
   i++
}

/*
   * FUNÇÃO PARA PESQUISAR E RETORNAR SOMENTE A COLUNA
*/
GS_SearchColumns(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   planilha := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   ; y célula da coluna header (id, nome, categoria, url) , x = linha
   for x,y in strsplit(substr(planilha, 1, instr(planilha,"`r")-1),",")
      (VarPesquisarDados = SubStr(y, 2, -1)) && pos := X ; se o campo pesquisa for igual a alguma coluna, pos = grava a posicao da coluna, se é a 3º ou 4ª coluna...
   ; DELETAR TODAS COLUNAS
   Loop, % LV_GetCount("Column")
      LV_DeleteCol(1)
   ; DELETAR TODAS AS LINHAS
   LV_Delete()
   ; ADICIONAR SOMENTE 1 COLUNA, QUE É A COLUNA PESQUISADA
   LV_InsertCol(1, , VarPesquisarDados)
   GuiControl, -Redraw, LVAll
   ; msgbox % pos
   for x,y in strsplit(planilha,"`n","`r")
      loop, parse, y, CSV
         if (x>1 && a_index = pos)
            LV_add("",a_loopfield)
   SB_SetText(LV_GetCount() " match(es)")
   LV_ModifyCol(1,"AutoHdr")
   ; SE A PESQUISA DE COLUNA RETORNAR NADA (0) - ATUALIZAR A PLANILHA
   If(LV_GetCount() = 0){
      GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   }
   GuiControl, +Redraw, LVAll
   i++
}

/*
   * FUNÇÃO PARA ATUALIZAR PLANILHA, RESET NA PLANILHA
*/
GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   LV_Delete()
   Gui Submit, NoHide
   GS_GetCSV_ToListView(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
}
/*
   * FUNÇÃO PARA REMOVER DADOS DUPLICADOS DE UM ARRAY
*/
RmvDuplic(object) {
   secondobject:=[]
   Loop % object.Length()
      {
         value:=Object.RemoveAt(1) ; otherwise Object.Pop() a little faster, but would not keep the original order
      Loop % secondobject.Length()
         If (value=secondobject[A_Index])
             Continue 2 ; jump to the top of the outer loop, we found a duplicate, discard it and move on
      secondobject.Push(value)
   }
   Return secondobject
}
/*
   * FUNÇÃO PARA TRATAR O IDIOMA DA DOCUMENTAÇÃO
*/

/*
   * FUNÇÃO PARA CHECAR A URL DA PLANILHA SELECIONADA NO COMBOBOX DA GUI "ALTERAR CONFIGURAÇÕES"
*/
checkSpreadsheetLink(PlanilhaLink){
   /*
      IMPORTANTE:
      A COLUNA E DA PLANILHA PRECISA TER UMA FÓRMULA PARA GERAR O ARRAY DOS DADOS
   */
   Gui Submit, NoHide
   ; msgbox %templateDimensoes%

   ; atualizar a url do google sheet TEMPLATE 1
   if(PlanilhaLink = "Documentações Analytics")
      {
         linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"
      }
      ; TEMPLATE 2
      else if(PlanilhaLink = "Documentações Programação")
      {
         linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"         
      }
      else if(PlanilhaLink = "Cursos")
      {
         linkPlanilha := "https://docs.google.com/spreadsheets/d/1_flbbi427JI7NiIk4ZGZvAM9eRBM4dd_gTDFgw3Npo8/edit#gid=0"
      }
      else if(PlanilhaLink = "Relatórios")
      {
         linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"
      }
      else if(PlanilhaLink = "Outros")
      {
         linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"
      }
      ; TRATAR PELA URL DA PLANILHA
      Else{
          if(RegExMatch(PlanilhaLink, "i).*docs.google.com/.+\/d\/.+\/")){
            RegExMatch(PlanilhaLink, "i).*\/d\/.+\/", UrlCode) 
            ; aceitar e usar o iniread iniread
          }else{
            MsgBox, 4112 , Erro na URL do Site!, URL Inválida`n- Copie e Cole uma URL do Google Sheets válida!
            ; Resetar/Limpar o valor do ComboBox
            GuiControl,ConfigFile:Choose, PlanilhaLink, ""
          }
          
      }
      Return PlanilhaLink
}


/*
   *
   *
   *
   * LABELS
*/
/*
   * AO SELECIONAR UMA TAB, DEFINIR O BOTÃO PADRÃO
*/
TabLabel:
GuiControlGet, h_Tab,, TabVariable
; msgbox % h_Tab
If (h_Tab="GA4")
   {
       GuiControl, +Default, VarAbrirDoc1
   }
   Else If (h_Tab="All")
   {
       GuiControl, +Default, BtnPesquisar
   }
   Else If (h_Tab="GDS")
   {
       GuiControl, +Default, VarAbrirDoc2
   }
   Else If (h_Tab="BigQ")
   {
       GuiControl, +Default, VarAbrirDoc3
   }
   Else If (h_Tab="Pixels")
   {
       GuiControl, +Default, VarAbrirDoc4
   }
   Else If (h_Tab="GTM")
   {
       GuiControl, +Default, VarAbrirDoc5
   }
Return
AbrirDoc:
Return

RecuperarPlanilha:
   Gui Submit, NoHide
   RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   ; msgbox % capture_sheetURL_key1
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   ; msgbox % capture_sheetURL_name1
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" sheetURL_key "/gviz/tq?tqx=out:" sheetURL_format "&range=" sheetURL_range "&sheet=" sheetURL_name "&tq=" GS_EncodeDecodeURI(sheetURL_SQLQuery)
   ; msgbox %PlanilhaTipoExportacao% %PlanilhaLink% %PlanilhaNomeId% %PlanilhaRange% %PlanilhaQuery%
   sheetData_All := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
Return
/*

*/
ValidarLink:
Gui Submit, NoHide
checkSpreadsheetLink(PlanilhaLink)
Return

AbrirCurso:
Gui, Submit, NoHide
AHK_GetControls()
Return

ListViewListener:
Gui Submit, NoHide
if(CheckIdiomaPt)
   GS_GetListView_Click("?hl=pt-br")
Else
   GS_GetListView_Click("?hl=en")
Return

; LABEL PARA CAPTURAR O CLIQUE NO BOTÃO ATUALIZAR LISTA
AtualizarPlanilha:
   GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
Return

/*
* LABELS DO MENU BAR
***
*/

MenuAcoesApp:
If(InStr(A_ThisMenuItem, "Sair"))
   ExitApp
Else If(InStr(A_ThisMenuItem, "Reiniciar"))
   Reload
return
/*

*/

MenuEditarBase:
   If(InStr(A_ThisMenuItem, "Trocar Planilha e/ou Configurações"))
   {
       ; ^n::
  ; MsgBox, Open Menu was clicked
  Gui, ConfigFile:Font, S11
  Gui, ConfigFile:New, +AlwaysOnTop -Resize -MinimizeBox -MaximizeBox, Alterar Configurações da Planilha
  /*
      * COLUNA 1
  */
  Gui, ConfigFile:Add, Text,center h20 +0x200, Alterar Link da Planilha:
  IniRead, PlanilhaLink, %iniPath%, planilha, linkPlanilha
  Gui ConfigFile:Add, ComboBox, y+5 w415 center vPlanilhaLink hwndDimensoesID gValidarLink,Documentações Analytics|Documentações Programação|Cursos|Relatórios|%PlanilhaLink%

  Gui, ConfigFile:Add, Text, center h20 +0x200, Nome/ID da aba da Planilha(Worksheet)
  Gui, ConfigFile:Add, Edit, w415 y+5 vPlanilhaNomeId

  /*
      * COLUNA 2
  */
  Gui, ConfigFile:Add, Text,section center h20 +0x200, Tipo de Exportação:
  Gui, ConfigFile:Add, ComboBox, vPlanilhaTipoExportacao w100 hwndCursosIDAll y+5 w200 center, CSV||HTML|JSON
  Gui, ConfigFile:Add, Text, ys x+10 center h20 +0x200, Range de Dados:
  Gui, ConfigFile:Add, Edit, vPlanilhaRange w205 y+5
  /*
      * FORA DAS COLUNAS
  */
  
  Gui, ConfigFile:Add, Text, xs y+10 center h20 +0x200, Query: 
  Gui, ConfigFile:Add, Edit, vPlanilhaQuery w420 y+5 r2,

  gui, font, S13 ;Change font size to 12
  gui, ConfigFile:Add, Button, center y+15 w100 h25 Default gSaveToIniFile, &Salvar
  Gui, ConfigFile:Show, xCenter yCenter
  ControlFocus, Edit1, Cadastrar Nova Doc - Felipe Lullio
  Gosub, ReadIniFile
   }
   Else If(InStr(A_ThisMenuItem, "trocar planilha(aba)"))
      run x
   Else If(InStr(A_ThisMenuItem, "alterar formato de exporta"))
      Run x
   Else If(InStr(A_ThisMenuItem, "alterar range"))
      Run x
   Else If(InStr(A_ThisMenuItem, "query"))
      Run x
Return

MenuAbrirLink:
; MsgBox, %A_ThisMenuItem%
If(InStr(A_ThisMenuItem, "abrir planilha"))
   Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit?usp=sharing"
Else If(InStr(A_ThisMenuItem, "Pasta Drive"))
   Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "https://drive.google.com/drive/folders/1m9rlPqx710icPobioyCU4FrcswwVGsdI?usp=sharing"
Else If(InStr(A_ThisMenuItem, "Pasta Script"))
   Run, %A_ScriptDir%

If(InStr(A_ThisMenuItem, "cursos udemy"))
   Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "https://www.udemy.com/home/my-courses/lists/"
Else If(InStr(A_ThisMenuItem, "Desenvolvedor"))
   Run, https://lullio.com.br
Else If(InStr(A_ThisMenuItem, "Sobre o programa"))
{
   Run, https://projetos.lullio.com.br/control-video-study
   Run, https://github.com/lullio/ahk-chrome-control-videos
}
Else If(InStr(A_ThisMenuItem, "WhatsApp"))
   Run, https://wa.me/5511991486309
return

/*
--------------------------
--------------------------
*/
/*
TRATAMENTO DA STATUS BAR
*/

/*
   AO CLICAR EM UMA POSIÇÃO DA STATUSBAR
*/
StatusBarLinks:
Gui Submit, Nohide
   ; msgbox %MyStatusBar%
   ; msgbox %A_EventInfo%
   ; if(A_GuiEvent == "Normal"){
   ;    msgbox %A_EventInfo%
   ; }
   ; recarregar página
   If(A_GuiEvent == "Normal" && A_EventInfo == 1){
   ; trazer página para frente
      
   }Else If(A_GuiEvent == "Normal" && A_EventInfo == 2){
      
   }Else If(A_GuiEvent == "Normal" && A_EventInfo == 3){
      Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit?usp=sharing"
   }Else If(A_GuiEvent == "Normal" && A_EventInfo == 4){
   }
Return

PesquisarDados:
Gui Submit, NoHide
If(CheckPesquisarColuna = true){ ; se o checkbox estiver marcado
   GS_SearchColumns(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
}else{
   GS_SearchRows(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
}
Return

ListenerIdioma:
if(CheckIdiomaPt = 1)
 idioma := "?hl=pt-br"
Else If (CheckIdiomaPt = 0)
 idioma := "?hl=en"

Return

