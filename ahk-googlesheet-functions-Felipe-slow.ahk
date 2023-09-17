/*
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
SkinForm(DLLPath,Apply, StylesPath "\MacLion3.msstyles") ; cosmo. msstyles
SkinForm(DLLPath,Param1 = "Apply", SkinName = ""){
  if(Param1 = Apply){
    DllCall("LoadLibrary", str, DLLPath)
    DllCall(DLLPath . "\USkinInit", Int,0, Int,0, AStr, SkinName)
  }else if(Param1 = 0){
    DllCall(DLLPath . "\USkinExit")
  }
}
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
; Run, %iniPath%

Gui, Destroy
Gui,+AlwaysOnTop ; +Owner
gui, font, S11 ;Change font size to 12
/*
MENU BAR
*/
Menu, FileMenu, Add, &Abrir Planilha Atual`tCtrl+O, MenuAbrirLink
Menu, FileMenu, Add ; with no more options, this is a seperator
Menu, FileMenu, Add, &Abrir Pasta Automations Drive`tCtrl+A, MenuAbrirLink
Menu, FileMenu, Add, &Abrir Pasta Documentações Template Drive`tCtrl+D, MenuAbrirLink
Menu, FileMenu, Add, &Abrir Pasta Documentações Oficiais Drive, MenuAbrirLink
Menu, FileMenu, Add, &Abrir Pasta Pixels Drive`tCtrl+P, MenuAbrirLink
Menu, FileMenu, Add, &Abrir Pasta do Script`tCtrl+E, MenuAbrirLink

Menu, EditMenu, Add, Trocar Planilha e suas Configurações`tCtrl+S, MenuEditarBase
Menu, EditMenu, Add, Colunas e Pesquisas`tCtrl+P, MenuEditarBase
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
Gui, Add, Statusbar, gStatusBarLinks vMyStatusBar,

/*
   * EDITAR TEXTO DA STATUS BAR
*/
SB_SetParts(120, 100, 100)
SB_SetText("Abrir Planilha", 2)
SB_SetText("Gerar Tabs", 3)
Gui Font, S9
global countGerarTabs:=1
/*
   * DEFINIR TODAS AS TABS
*/
Gui, Add, Tab3, vTabVariable hwndhwnd, All ; |GA4|GDS|BigQ|Pixels|GTM
Gui Font, S10
; CRIAR A PRIMEIRA TAB
Gui Tab, All 

/*
 * ********* TAB 1
*/
; se quiser que apareça nome do grupo tirar o -Hdr
Gui, Add, ListView, r15 Grid NoSortHdr vLVAll w450 gListViewListener,
Gui, Add, Edit, h29 vVarPesquisarDados w230 y+10 section, GA4
Gui, Add, Button, vBtnPesquisar x+10 w100 h30 gPesquisarDados Default, Pesquisar
Gui, Add, Button, vBtnAtualizar x+10 w100 h30 gAtualizarPlanilha, Atualizar
; Gui, Add, Button, vBtnAtualizar1 y+5 w100 h30 gGerarTabsListas, Gerar Tabs
; Gui, Add, Checkbox, vCheckIdiomaPt Checked1 xs y+10, abrir documentação em português
Gui, Add, Checkbox, vCheckPesquisarColuna Checked0 xs ys+35, pesquisar por coluna
/*
   * FORA DA TAB
*/
Gui, Tab
Gui, Add, Checkbox, vCheckIdiomaPt Checked1 xs y+5 center, abrir documentação em português
/*
   O RESTO DAS TABS É GERADO DINÂMICAMENTE COM BASE NOS DADOS DA PLANILHA
*/
Gui, Show, AutoSize , Web Analytics Links Helper - Felipe Lullio
GuiControl, +Default, BtnPesquisar ; Definir o botão Pesquisar como Padrão
ControlFocus, Edit1, Web Analytics Links ; Dar foco no input Edit de Pesquisa

; Gui, ListView, LVAll
/*
   *VARIÁVEIS PARA FORMAR A URL DO GOOGLE SHEET*
   - Somente a sheetURL_key é obrigatória
     
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" sheetURL_key "gviz/tq?tqx=out:" sheetURL_format "&range=" sheetURL_range "&sheet=" sheetURL_name "&tq=" sheetURL_SQLQueryEncoded
   msgbox % fullSheetURL 
*/
; sheetURL_key := "1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34" ; id da pasta de trabalho/arquivo
; sheetURL_name := "All-Docs" ; nome ou id da aba / guia / planilha
; sheetURL_format := "csv" ; csv, html ou json
; sheetURL_range := "" ; A1:C99
; sheetURL_SQLQueryGA4Doc := "select * where D matches '^GA4.*' AND D is not null"
; sheetURL_SQLQuery := "select * where A matches '.*' AND A is not null"
; sheetURL_SQLQueryEncoded = % GS_EncodeDecodeURI(sheetURL_SQLQuery)
; global i:=1 ; contas quantas vezes clicou no botão (botão Pesquisar)
; global Colunas := [] ; salvar os nomes das colunas pela function GS_GetCSV_ToListView()
/*
   * LER ARQUIVO DE CONFIGURAÇÃO
*/
ReadIniFile:
Gui Submit, NoHide
global PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
; If(!PlanilhaLink)
;    msgbox hi
; Link da Planilha
IniRead, PlanilhaLink, %iniPath%, planilha, linkPlanilha
GuiControl, ConfigFile:Choose, PlanilhaLink, %PlanilhaLink%
; Tipo de Exportação
IniRead, PlanilhaTipoExportacao, %iniPath%, planilha, tipoExportacao
GuiControl, ConfigFile:Choose, PlanilhaTipoExportacao, %PlanilhaTipoExportacao%
; Aba da Planilha
IniRead, PlanilhaNomeId, %iniPath%, planilha, abaPlanilha
GuiControl, ConfigFile:Text, PlanilhaNomeId, %PlanilhaNomeId%
; Regex Nome
IniRead, PlanilhaRegexNome, %iniPath%, planilha, regexNomePlanilha
GuiControl, ConfigFile:Text, PlanilhaRegexNome, %PlanilhaRegexNome%
; Regex URL
IniRead, PlanilhaRegexURL, %iniPath%, planilha, regexURLPlanilha
GuiControl, ConfigFile:Text, PlanilhaRegexURL, %PlanilhaRegexURL%
; Range de Dados
IniRead, PlanilhaRange, %iniPath%, planilha, rangePlanilha
GuiControl, ConfigFile:Text, PlanilhaRange, %PlanilhaRange%
; Query
IniRead, PlanilhaQuery, %iniPath%, planilha, queryPlanilha
GuiControl, ConfigFile:Text, PlanilhaQuery, %PlanilhaQuery%
If(PlanilhaLink)
{
   GS_GetCSV_ToListView(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   global posicaoColunaNome := GS_GetCSV_Column(, ".*Nome.*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).ColumnPosition
   global posicaoColunaURL := GS_GetCSV_Column(, "i).*(URL|Site|link).*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).ColumnPosition
   global planilha := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   ; Run %iniFile%
}
Return
/*
   * ESCREVER NO ARQUIVO DE CONFIGURAÇÃO
*/
SaveToIniFile:
Gui Submit
 ; Link da Planilha
IniWrite, %PlanilhaLink%, %iniPath%, planilha, linkPlanilha
 ; Tipo de Exportação
IniWrite, %PlanilhaTipoExportacao%, %iniPath%, planilha, tipoExportacao
 ; Nome/ID da Aba
IniWrite, %PlanilhaNomeId%, %iniPath%, planilha, abaPlanilha
 ; Regex Nome
IniWrite, %PlanilhaRegexNome%, %iniPath%, planilha, regexNomePlanilha
 ; Regex URL
IniWrite, %PlanilhaRegexURL%, %iniPath%, planilha, regexURLPlanilha
 ; Range da Planilha
IniWrite, %PlanilhaRange%, %iniPath%, planilha, rangePlanilha
 ; Query da Planilha
IniWrite, %PlanilhaQuery%, %iniPath%, planilha, queryPlanilha
Notify().AddWindow("Configuração atualizada!`nClique no botão Atualizar para atualizar os dados!",{Time:5000,Icon:177,Background:"0x039018",Title:"SUCESSO",TitleColor:"0xF0F8F1", TitleSize:15, Size:15, Color: "0xF0F8F1"},"","setPosBR")
global posicaoColunaNome := GS_GetCSV_Column(, ".*Nome.*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).ColumnPosition
global posicaoColunaURL := GS_GetCSV_Column(, "i).*(URL|Site|link).*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).ColumnPosition
global planilha := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
checkSpreadsheetLink(PlanilhaLink)
; Run %iniFile%
Return

/* 
   * FUNÇÃO PARA O TRATAMENTO DOS DROPDOWN, PARA QUANDO VC ESCREVER O NOME DO CURSO JÁ PREENCHER O CURSO AUTOMATICAMENTE NO DROPDOWN
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
   PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
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
   ;  PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
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
   PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   ; msgbox % capture_sheetURL_key1
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   ; msgbox % capture_sheetURL_name1
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key1 "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&sheet=" capture_sheetURL_name1 "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)
   ; msgbox %PlanilhaTipoExportacao% %PlanilhaLink% %PlanilhaNomeId% %PlanilhaRange% %PlanilhaQuery%
   sheetData_All := GS_GetCSV(capture_sheetURL_key1, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, capture_sheetURL_name1)
   
   ; msgbox % sheetData_All
      
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
              ; salvar todos os nomes das colunas / header column
              Colunas.Push(SubStr(StrSplit(A_LoopField,",")[A_Index], 2, -1))
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

       LV_ModifyCol(1, "30 Right") 
       LV_ModifyCol(2, "250 Left") 
       LV_ModifyCol(3, "200 Left") 
       
       ; total de linhas
       TotalLinhas:
         totalLines := LV_GetCount()
         GuiControl, , TotalLinhas, Total de Linhas: %totalLines%
         SB_SetText("Total de Linhas: " totalLines, 1)
       Return {nomesColunas: coco, colunasHeader: [ColunaHeader1, ColunaHeader2, ColunaHeader3, ColunaHeader4, ColunaHeader5, ColunaHeader6, ColunaHeader7, ColunaHeader8, ColunaHeader9, ColunaHeader10, ColunaHeader11, ColunaHeader12, ColunaHeader13], Colunas: Colunas}
}

/*
   * FUNÇÃO PARA CAPTURAR AÇÃO AO CLICAR NA LISTVIEW
*/
GS_GetListView_Click(idioma, PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId, regexFindColumnName:= ".*Nome.*", regexFindColumnURL := "i).*(URL|Link).*", action := "openLink"){
   Gui Submit, NoHide
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
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
      ; msgbox % URLSanitized
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

; LABEL PARA GERAR AS TABS E COMBOBOX
GerarTabsListas:
   Gui Submit, NoHide
   if(countGerarTabs = 1){
      GerarTabsComboBox(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   }else{
      MsgBox, 4112 , Erro!, Não é possível gerar tabs repetidas!`n Reinicie o programa(CTRL+R), 8
   }
   countGerarTabs++
Return
/*
   * FUNÇÃO PARA CRIAR TODAS AS TABS E LISTAS COMBOBOX DE FORMA DINÂMICA
   * ESTOU CAPTURANDO A POSIÇÃO DA COLUNA "URL" E POSIÇÃO DA COLUNA "NOME", LÁ EM CIMA, COMO VARIÁVEL GLOBAL
*/
GerarTabsComboBox(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   ; msgbox % posicaoColunaNome ; variável global
   sheetData_All := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId) ; Select * limit 1
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   ColumnCategory := GS_GetCSV_Column(, "i)Categoria",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).arrColumnSanitize ; ColumnData.variavelJavascript ColumnData.arrColumn ColumnData.strColumn
   ; remover itens duplicados, aqui são as categorias exibidas na planina , ex: BQ_GA4_Export 
   UniqueColumnCategory := RmvDuplic(ColumnCategory)
   quantidadeColunasPlanilha := UniqueColumnCategory.MaxIndex() ; QUANTIDADE DE CATEGORIAS NA PLANILHA
   ArrTabs := []
   Categorias :=  []
   countComboBoxColuna1 := 0
   countComboBoxColuna2 := 0
   aspa := """"
   ; remover o primeiro item do array pois é a primeira linha, que é o termo "Categoria"
   ; UniqueColumnCategory.RemoveAt(1)
   ; msgbox % UniqueColumnCategory.MaxIndex()
   ; Loop, % UniqueColumnCategory.MaxIndex()
   /*
      * DEFINIR QUAIS SERÃO AS TABS DA GUI COM BASE NA COLUNA CATEGORIRAS, PEGANDO SOMENTE O COMEÇO
      * ADICIONAR SOMENTE A PRIMEIRA PARTE DA CATEGORIA QUE TÁ NA PLANILHA, EX: GA4_ GA3_
   */
   For index, categoria in UniqueColumnCategory
   {
      ; VALIDAR PARA ADICIONAR PALAVRAS IGUAIS, ADICIONAR SOMENTE SE O ITEM EXIBIDO FOR DIFERENTE DO ITEM ANTERIOR
      If(Index != 1 && StrSplit(UniqueColumnCategory[index], "_")[1] != StrSplit(UniqueColumnCategory[index-1], "_")[1])
      {
         ArrTabs.push(StrSplit(categoria, "_")[1])
      }
   }
   ; Loop % UniqueColumnCategory.MaxIndex()
   ;    msgbox % UniqueColumnCategory[A_Index]
   
   ; Loop, % Colunas.MaxIndex()
   ; {
   ;    Coluna%A_Index%%Colunas[A_Index]% = "asd"
   ;    msgbox % Coluna%A_Index%%Colunas[A_Index]%
   ; }


   ; sheetData_ColumnDataArr := []
   ; sheetData_ColumnDataArrSanitize := []
   ; sheetData_ColumnDataStr := ""
   ; sheetData_ColumnDataStrSanitize := ""
   ; sheetData_ColumnPosition := 0
   ; sheetData_ColumnName := ""
   ; sheetData_ColumnPosition := ""
   ;  regexFindColumn := "i)Categoria"

   ; for key, category in UniqueColumnCategory
   ; {
      ; CRIAR AS TABS (VAI SER O INICIO DO TEXTO DA CATEGORIA DA PLANILHA)
      For, index, tabCategory in ArrTabs 
      {
         ; CRIAR AS TABS - CÓDIGO TOP
         GuiControl,, TabVariable, %tabCategory%
         ; ATIVAR A TAB, OS CÓDIGOS ABAIXO DESSE CÓDIGO SERÃO APLICADOS A TAB SELECIONADA/ATIVADA
         Gui, Tab, %tabCategory%
         ; CRIAR UM TEXTO VAZIO PARA ELE SER A COLUNA 1 = A SEÇÃO, O ÚNICO QUE VAI TER O "SECTION" PARA DEFINIR A COLUNA 1
         Gui Add, Text, section x+15 y+10 center, TAB - %tabCategory%
         ; CONTAR O NÚMERO DE COMBOBOX EXISTENTE NA TAB, SE PASSAR DE 6, CRIAR NOVA COLUNA
         countComboBoxColuna1 := 0
         countComboBoxColuna2 := 0
      ; PARA CADA CATEGORIA EXISTENTE (BQ_GA4_EXPORT, ....)
      For, key, category in UniqueColumnCategory
      {
               ; msgbox %category%
               Coluna%category% := []
               ; ! TIPOS DE EVENTOS GA4
               ; dropdown 1 - principais cursos
               /*
               -----------
               ----------- 1ª COLUNA
               -----------
               */
               ; somente se a categoria tem o mesmo começo do nome da tab, ex: categoira GA4_Tipos_de_Eventos , tab: GA4
               If(InStr(category, tabCategory) && countComboBoxColuna1 <=4)
               {
                  Gui Add, Text, y+10, %category%
                  Loop, parse, sheetData_All, `n ; PROCESSAR CADA LINHA DA TABELA/PLANILHA
                     {
                        LineNumber := A_Index ; Index da linha
                        LineContent := A_LoopField ; Conteúdo da linha, todos valores da linha, a 1ª linha vai ser o HEADER(vc consegue capturar os headers das colunas)
                        if(InStr(LineContent, category))
                        {
                           Coluna%category%.push(StrSplit(LineContent, ",")[posicaoColunaNome])
                           ; LISTA DE NOMES PERTENCENTES A CATEGORIA, VAI SER USADA 
                           ListaDeNomes .= RegExReplace(StrSplit(LineContent, ",")[posicaoColunaNome] "|", aspa , "")
                           ; GuiControlGet, var,, % AXID
                           ; GuiControl,,  "|hello"
                        }
                     }
                     Gui, Add, ComboBox, y+5 w200 hwndIdEventos, %ListaDeNomes%
                     ; Resetar a lista de nomes após concluir de preencher um COMBOBOX
                     ListaDeNomes := ""
                     countComboBoxColuna1++
               /*
                  * CRIAR UMA SEGUNDA COLUNA QUANDO PASSAR DE 3 COMBOBOXES/LISTAS
               */
               }else if(InStr(category, tabCategory) && countComboBoxColuna2 = 0){ ; se passou de 6 combobox na tab
                  Gui, Add, Text, ys+25 x+35, %category%
                  Loop, parse, sheetData_All, `n ; PROCESSAR CADA LINHA DA TABELA/PLANILHA
                     {
                        LineNumber := A_Index ; Index da linha
                        LineContent := A_LoopField ; Conteúdo da linha, todos valores da linha, a 1ª linha vai ser o HEADER(vc consegue capturar os headers das colunas)
                        if(InStr(LineContent, category))
                        {
                           Coluna%category%.push(StrSplit(LineContent, ",")[posicaoColunaNome])
                           ; LISTA DE NOMES PERTENCENTES A CATEGORIA, VAI SER USADA 
                           ListaDeNomes .= RegExReplace(StrSplit(LineContent, ",")[posicaoColunaNome] "|", aspa , "")
                           ; GuiControlGet, var,, % AXID
                           ; GuiControl,,  "|hello"
                        }
                     }
                     countComboBoxColuna2++ ; PARA SABER SE JÁ CRIOU UM "YS", se não vai ficar criando colunas
                     Gui, Add, ComboBox, y+5 w200 hwndIdEventos, %ListaDeNomes%
                     ; Resetar a lista de nomes após concluir de preencher um COMBOBOX
                     ListaDeNomes := ""
               /*
                  * APÓS TER CRIADO A SEGUNDA COLUNA, NÃO DEVE INSERIR MAIS COLUNAS, INSERIR UM ABAIXO DO OUTRO
               */
               }else if(InStr(category, tabCategory)){ ; JÁ CRIOU UMA SEGUNDA COLUNA, ENTÃO AGORA DEVE INSERIR OS COMBOBOX UM ABAIXO DO OUTRO
                  Gui, Add, Text, y+10 , %category%
                  Loop, parse, sheetData_All, `n ; PROCESSAR CADA LINHA DA TABELA/PLANILHA
                     {
                        LineNumber := A_Index ; Index da linha
                        LineContent := A_LoopField ; Conteúdo da linha, todos valores da linha, a 1ª linha vai ser o HEADER(vc consegue capturar os headers das colunas)
                        if(InStr(LineContent, category))
                        {
                           Coluna%category%.push(StrSplit(LineContent, ",")[posicaoColunaNome])
                           ; LISTA DE NOMES PERTENCENTES A CATEGORIA, VAI SER USADA 
                           ListaDeNomes .= RegExReplace(StrSplit(LineContent, ",")[posicaoColunaNome] "|", aspa , "")
                           ; GuiControlGet, var,, % AXID
                           ; GuiControl,,  "|hello"
                        }
                     }
                  Gui, Add, ComboBox, y+5 w200 hwndIdEventos, %ListaDeNomes%
                  countComboBoxColuna2++
                  ; Resetar a lista de nomes após concluir de preencher um COMBOBOX
                  ListaDeNomes := ""                  
               }
         ;   Gui, Show
     } ; FIM DO LOOP DAS CATEGORIAS ÚNICAS DA PLANILHA
     /*
         * AO TERMINAR DE GERAR TODOS OS COMBOXES E ESTAR CHEGANDO PERTO DE CRIAR OUTRA TAB
     */
   /*
     * RESOLVER O PROBLEMA DE O ÚLITMO CONTROL SER CONSIDERADO O CONTROL DA COLUNA2, COMO O CONTROL FICA BEM ANTES DO ÚLTIMO DA COLUNA 1, OS ELEMENTOS ABAIXO COMO OS BOTÕES SÃO INSERIDOS BEM PRA CIMA FICANDO ERRADO
   */
   ; * total de combobox na coluna 1 e 2
   ; msgbox % "total coluna1: " countComboBoxColuna1 "totalColuna2: " countComboBoxColuna2
   ; * CASO TENHA PELO MENOS UM COMBOBOX NA COLUNA 2 E O NÚMERO DE COMBOBOX NA COLUNA 2 É MENOR QUE O DA COLUNA1
     While(countComboBoxColuna2 < countComboBoxColuna1 && countComboBoxColuna2 > 0)
      {
         ; msgbox % countComboBoxColuna2
         countComboBoxColuna2++
         Gui, Add, Text, +Hidden y+10 
         Gui, Add, ComboBox, +Hidden y+5 w200 hwndIdEventos
      }
     /*
         * BOTÕES DA TAB
     */
      Gui, Add, Link, xs y+20, <a>Root-Doc</a> | <a>What's New</a> | <a>Blog</a> | <a>Notion</a>
      ; Botões
      gui, font, S11
      gui, Add, Button, y+10 w200 gAbrirDoc Default, &Abrir Doc
      gui, Add, Button, w150 x+20 Cancel gCancel, &Cancelar
      ; Gui, Add, Checkbox, Checked1 x+10, pt-br?
   } ; FIM DO LOOP DA TAB
   ;   for i, val in Categorias
   ;  {
   ;     msgbox %val%
   ;     for in, valn in Coluna%val%
   ;        msgbox %valn%
   ;  }
   ;  AHK_GetControls()
 Return 
} 

/*
   * FUNÇÃO PARA CAPTURAR TODOS OS CONTROLS DA GUI
   * * CAPTURAR SOMENTE OS COMBOBOX COM "IF STATEMENT"
   * * LOOP 1: Loop em todos controls ahk
   * * * LOOP 2: Se for um combobox rodar um Loop de todas as linhas da planilha e comparar com o texto que está no combobox
*/
AHK_GetControls(idioma, PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId,searchControls := "ComboBox"){
   Gui, Submit, NoHide
   aspa := """"
   ; msgbox % idioma
   ; msgbox % posicaoColunaNome ; variável global
   ; msgbox % posicaoColunaURL ; variável global
   ; * CAPTURAR TODOS OS CONTROLS DA GUI
   WinGet, ActiveControlList, ControlList, A
      ; Loop, Parse, ActiveControlList, `n
      /*
         * LOOP EM TODOS CONTROLS DA GUI
      */
      For index, control in StrSplit(ActiveControlList, "`n")
         {
            ; CAPTURAR O TEXTO DE CADA CONTROL
            ControlGetText, TextoDoControl, %control%
            ; SALVAR OS DADOS DO CONTROL EM UM ARQUIVO
            FileAppend, %index%`t%control%`t%TextoDoControl%`n, C:\Controls.txt
            ; DAR FOCO NO CONTROL
            GuiControl, Focus, %control%
            ; RETORNAR O NOME/VARIÁVEL DO CONTROL QUE ESTÁ COM FOCO
            GuiControlGet,varName, FocusV
            /*
               * CAPTURANDO SOMENTES OS ComboBoXES
            */               
            if(InStr(control, searchControls)) ; * se for um combobox
            {
               for index,row in strsplit(planilha,"`n","`r") ; * LOOP EM CADA LINHA DA PLANILHA
                  ; * SE O TEXTO DO CONTROL ESTIVER DENTRO DO TEXTO DA LINHA DA PLANILHA
                  if (varName != "" && InStr(row, trim(varName))) ; * VARIÁVEL NÃO PODE ESTAR VAZIA, SE NÃO VAI ACABAR DANDO MATCH EM TODOS CONTROLS POR ESTAREM VAZIO SEM NADA SELECIONADO
                     {
                        ; msgbox % varName
                        ; msgbox % row
                        ; msgbox % InStr(row, varName)
                        ; RegExReplace(StrSplit(LineContent, ",")[posicaoColunaNome] "|", aspa , "")
                              /*
                        * ABRIR OS LINKS/URLS/DOCUMENTAÇÕES NO NAVEGADOR
                        ! IMPORTANTE: Caso tenha mais de um link na coluna, transformar em um array e fazer um loop para abrir os links
                        */
                        URLSemAspa := RegExReplace(StrSplit(row, ",")[posicaoColunaURL], aspa, "")
                        URLSanitized := StrReplace(URLSemAspa, "%idiomapt%", idioma)
                        ; msgbox % URLSanitized
                        For Index, URL in StrSplit(URLSanitized, " | ")
                           {
                              Run, %URL%
                           } 
                        varName := ""
                     }
            }
            if(InStr(control, "checkbox"))
               msgbox hi
         }
}
/*
   * FUNÇÃO PARA PESQUISAR E RETORNAR TODAS LINHAS E COLUNAS
*/
GS_SearchRows(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   cnt := 0
   Gui Submit, NoHide
   planilha := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   ; msgbox % planilha
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
   SB_SetText("Match(es) da última Pesquisa: " cnt,  4)
   loop, % lv_getcount("col")
      LV_ModifyCol(a_index,"AutoHdr")
   GuiControl, +Redraw, LVAll
   i++
   If(LV_GetCount() = 0){
      MsgBox, 4112 , Erro!, A Pesquisa não retornou nada`nVamos atualizar os dados!, 2
      GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
      ; Sleep, 500
      ; Notify().AddWindow("Erro",{Time:3000,Icon:28,Background:"0x990000",Title:"ERRO",TitleSize:15, Size:15, Color: "0xCDA089", TitleColor: "0xE1B9A4"},"w330 h30","setPosBR")
   }
}

/*
   * FUNÇÃO PARA PESQUISAR E RETORNAR SOMENTE A COLUNA
*/
GS_SearchColumns(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
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
      MsgBox, 4112 , Erro!, A Pesquisa não retornou nada`nVamos atualizar os dados!, 2
      GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
      ; Sleep, 500
      ; Notify().AddWindow("Erro",{Time:3000,Icon:28,Background:"0x990000",Title:"ERRO",TitleSize:15, Size:15, Color: "0xCDA089", TitleColor: "0xE1B9A4"},"w330 h30","setPosBR")
   }
   GuiControl, +Redraw, LVAll
   i++
}

/*
   * FUNÇÃO PARA ATUALIZAR PLANILHA, RESET NA PLANILHA
*/
GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId){
   PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   ; msgbox % capture_sheetURL_key1
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   ; msgbox % capture_sheetURL_name1
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key1 "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&sheet=" capture_sheetURL_name1 "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)
   ; msgbox %PlanilhaTipoExportacao% %PlanilhaLink% %PlanilhaNomeId% %PlanilhaRange% %PlanilhaQuery%
   sheetData_All := GS_GetCSV(capture_sheetURL_key1, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, capture_sheetURL_name1)
   PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   LV_Delete() ; deletar todas as linhas
   ; deletar todas as colunas
   Loop, % LV_GetCount("Column") 
      LV_DeleteCol(1)
   ; executar a planilha novamente
   GS_GetCSV_ToListView(capture_sheetURL_key1, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, capture_sheetURL_name1)
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
   * FUNÇÃO PARA CHECAR A URL DA PLANILHA SELECIONADA NO COMBOBOX DA GUI "ALTERAR CONFIGURAÇÕES"
*/
checkSpreadsheetLink(PlanilhaLink){
   /*
      IMPORTANTE:
      A COLUNA E DA PLANILHA PRECISA TER UMA FÓRMULA PARA GERAR O ARRAY DOS DADOS
   */
   ; msgbox %templateDimensoes%

   ; atualizar a url do google sheet TEMPLATE 1
      if(PlanilhaLink = "Documentações Analytics")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"
      ; TEMPLATE 2
      else if(PlanilhaLink = "Documentações Banco de Dados")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1ZmlzAhTGDPCsAz9yHAQGHEGPFdLDh1sCE6D7ePHNLjM/edit?usp=sharing"        
      ; TEMPLATE 2
      else if(PlanilhaLink = "Documentações Programação")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"        
      else if(PlanilhaLink = "Cursos")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1_flbbi427JI7NiIk4ZGZvAM9eRBM4dd_gTDFgw3Npo8/edit#gid=0"
      else if(PlanilhaLink = "Relatórios")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"
      else if(PlanilhaLink = "Documentações GAPS")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=218001466"
      else if(PlanilhaLink = "Outros")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"
      ; TRATAR PELA URL DA PLANILHA
      Else If(RegExMatch(PlanilhaLink, "i).*docs.google.com/.+\/d\/.+\/"))
         Return linkPlanilha := PlanilhaLink
      Else If(!InStr(PlanilhaLink, "https://docs.google.com/spreadsheets") || RegexMatch(PlanilhaLink, "\s{0,}"))
      {
         ; MsgBox, 4112 , Erro na URL do Site!, URL Inválida`n- Copie e Cole uma URL do Google Sheets válida!
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"
         ; GoSub, MenuEditarBase
         ; Resetar/Limpar o valor do ComboBox
         ; GuiControl,ConfigFile:Choose, PlanilhaLink, 1
      }
      Return linkPlanilha
}

/*
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * ************************************ ************************************ ***********************************
   * LABELS
*/
/*
   * AO SELECIONAR UMA TAB, DEFINIR O BOTÃO PADRÃO
*/
; TabLabel:
;    Gui Submit, NoHide
;    GuiControlGet, h_Tab,, TabVariable
;    If (h_Tab="GA4")
;    {
;       ; Gui, Destroy
;        GuiControl, +Default, VarAbrirDoc1
;       ;  Gui, Add, Text,, &First Name:
;       ;  Gui, Add, Edit
;       ;  GuiControl, 1:Hide, TabVariable
;       ;  GuiControl, 1:Hide, h_tab
;    }
;    Else If (h_Tab="All")
;    {
;        GuiControl, +Default, BtnPesquisar
;    }
;    Else If (h_Tab="GDS")
;    {
;        GuiControl, +Default, VarAbrirDoc2
;    }
;    Else If (h_Tab="BigQ")
;    {
;        GuiControl, +Default, VarAbrirDoc3
;    }
;    Else If (h_Tab="Pixels")
;    {
;        GuiControl, +Default, VarAbrirDoc4
;    }
;    Else If (h_Tab="GTM")
;    {
;        GuiControl, +Default, VarAbrirDoc5
;    }
; Return
/*
   * RECUPERAR OS DADOS DA PLANILHA
*/
RecuperarPlanilha:
   Gui Submit, NoHide
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   ; msgbox % capture_sheetURL_key1
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   ; msgbox % capture_sheetURL_name1
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" sheetURL_key "/gviz/tq?tqx=out:" sheetURL_format "&range=" sheetURL_range "&sheet=" sheetURL_name "&tq="     GS_EncodeDecodeURI(sheetURL_SQLQuery)
   ; msgbox %PlanilhaTipoExportacao% %PlanilhaLink% %PlanilhaNomeId% %PlanilhaRange% %PlanilhaQuery%
   sheetData_All := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
Return

/*
   * VALIDAR O LINK DA PLANILHA, CONVERTER A OPÇÃO SELECIONADA PARA URL
*/
ValidarLink:
Gui Submit, NoHide
checkSpreadsheetLink(PlanilhaLink)
Return

/*
   * AO CLICAR NO BOTÃO ABRIR DOC DAS TABS
*/
AbrirDoc:
   Gui, Submit, NoHide
   if(CheckIdiomaPt)
      AHK_GetControls("?hl=pt-br", PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   Else
      AHK_GetControls("?hl=en", PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
Return
/*
   * AO SELECIONAR UM ITEM NA LISTVIEW VAI CHAMAR A FUNÇÃO DE CLICAR E ANTES VAI TRATAR O IDIOMA ESCOLHIDO , CHECKBOX IDIOMA
*/
ListViewListener:
   Gui Submit, NoHide
   if(CheckIdiomaPt)
      GS_GetListView_Click("?hl=pt-br",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   Else
      GS_GetListView_Click("?hl=en",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
Return
/*
   * AO CLICAR NO BOTÃO "ATUALIZAR", VAI EXCLUIR A LISTVIEW E CRIAR NOVAMENTE
*/
AtualizarPlanilha:
   PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
Return
/*
   * CASO O CHECKBOX DE "PESQUISAR POR COLUNA" ESTEJA MARCADO
*/
PesquisarDados:
   Gui Submit, NoHide
   If(CheckPesquisarColuna = true){ ; se o checkbox estiver marcado
      GS_SearchColumns(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   }else{
      GS_SearchRows(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   }
Return

/*
   * LABELS DO MENU BAR
*/

MenuAcoesApp:
   If(InStr(A_ThisMenuItem, "Sair"))
      ExitApp
   Else If(InStr(A_ThisMenuItem, "Reiniciar"))
      Reload
return
MenuEditarBase:
   If(InStr(A_ThisMenuItem, "Trocar Planilha e suas Configurações"))
   {
       ; ^n::
  ; MsgBox, Open Menu was clicked
  Gui, ConfigFile:New, +AlwaysOnTop -Resize -MinimizeBox -MaximizeBox, Alterar Configurações da Planilha
  /*
  * COLUNA 1 - linha inteira
  */
  Gui, ConfigFile:Font, S10
  Gui, ConfigFile:Add, Text,center h20 +0x200, Alterar Link da Planilha:
  IniRead, PlanilhaLink, %iniPath%, planilha, linkPlanilha
  Gui ConfigFile:Add, ComboBox, y+5 w415 center vPlanilhaLink hwndDimensoesID gValidarLink,Documentações Analytics|Documentações Banco de Dados|Documentações Programação|Documentações GAPS|Cursos|Relatórios|%PlanilhaLink%

  Gui, ConfigFile:Add, Text, center h20 +0x200, Nome/ID da aba da Planilha(Worksheet)
  Gui, ConfigFile:Add, Edit, w415 y+5 vPlanilhaNomeId

  /*
      * COLUNA 2 - metade
  */
  Gui, ConfigFile:Add, Text,section center h20 +0x200, Regex Coluna Nome/Título
  Gui, ConfigFile:Add, Edit, vPlanilhaRegexNome hwndCursosIDAll y+5 w205 ,
  Gui, ConfigFile:Add, Text,center h20 +0x200, Tipo de Exportação:
  Gui, ConfigFile:Add, ComboBox, vPlanilhaTipoExportacao w100 hwndCursosIDAll y+5 w200 center, CSV||HTML|JSON
  /*
      * COLUNA 3 - metade
  */
  Gui, ConfigFile:Add, Text,ys x+10  h20 +0x200, Regex Coluna URL/Site (Ação)
  Gui, ConfigFile:Add, Edit, vPlanilhaRegexURL hwndCursosIDAll y+5 w205 ,
  Gui, ConfigFile:Add, Text, center h20 +0x200, Range de Dados:
  Gui, ConfigFile:Add, Edit, vPlanilhaRange w205 y+5
  /*
      * FORA DAS COLUNAS
  */
  Gui, ConfigFile:Font, S9
  Gui, ConfigFile:Add, Text, xs y+10 center h20 +0x200, Query: 
  Gui, ConfigFile:Add, Edit, vPlanilhaQuery w420 y+5 r2,
  Gui, ConfigFile:Font, S10

  gui, font, S13 ;Change font size to 12
  gui, ConfigFile:Add, Button, center y+15 w100 h25 Default gSaveToIniFile, &Salvar
  Gui, ConfigFile:Show, xCenter yCenter
  If(PlanilhaLink){
     Gosub, ReadIniFile
     Return
  }
   Return
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
   Gui Submit, NoHide
   ; MsgBox, %A_ThisMenuItem%
   If(InStr(A_ThisMenuItem, "Abrir Planilha"))
   {
      Run, % checkSpreadsheetLink(PlanilhaLink)
   }
   Else If(InStr(A_ThisMenuItem, "Abrir Pasta Pixels"))
      Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "https://drive.google.com/drive/folders/1m9rlPqx710icPobioyCU4FrcswwVGsdI?usp=sharing"
   Else If(InStr(A_ThisMenuItem, "Abrir Pasta Automations"))
      Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "https://drive.google.com/drive/folders/1f119V2GprzjM304wedJ9cXNocPro-G8s?usp=sharing"
   Else If(InStr(A_ThisMenuItem, "Abrir Pasta Documentações Template"))
      Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "https://drive.google.com/drive/folders/192gN6_efPFzvMX2lO6O_l1BKLd5vgUdV?usp=drive_link"
   Else If(InStr(A_ThisMenuItem, "Abrir Pasta Documentações Oficiais"))
      Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "https://drive.google.com/drive/folders/1gQzSELhAoqa00T8HQMg8iqTeXvtmuT-c?usp=drive_link"
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
   * STATUS BAR, TRATAR STATUSBAR
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
      Run, % checkSpreadsheetLink(PlanilhaLink)
   }Else If(A_GuiEvent == "Normal" && A_EventInfo == 3){
      GoSub, GerarTabsListas
   }Else If(A_GuiEvent == "Normal" && A_EventInfo == 4){
   }
Return
; ListenerIdioma:
; if(CheckIdiomaPt = 1)
;  idioma := "?hl=pt-br"
; Else If (CheckIdiomaPt = 0)
;  idioma := "?hl=en"

; Return

