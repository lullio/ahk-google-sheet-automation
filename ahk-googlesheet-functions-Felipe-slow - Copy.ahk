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


Gui, Add, Tab3, vTabVariable hwndhwnd, All ; |GA4|GDS|BigQ|Pixels|GTM
Gui Font, S10
; CRIAR A PRIMEIRA TAB
Gui Tab, All

/*
 * ********* TAB 1
*/
; se quiser que apareça nome do grupo tirar o -Hdr
Gui, Add, ListView, r15 Grid NoSortHdr vLVAll w460 gListViewListener ,
Gui Font, S12
Gui, Add, Edit, h29 vVarPesquisarDados w230 y+10 section cblue, .*ecommerce.*
Gui Font, S10,
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
GoSub, ReadIniFile
Gui, Show, AutoSize , Web Analytics Links Helper - Felipe Lullio
; Gui, +Resize
GuiControl, +Default, BtnPesquisar ; Definir o botão Pesquisar como Padrão
ControlFocus, Edit1, Web Analytics Links ; Dar foco no input Edit de Pesquisa
; GoSub, ReadIniFile ; aqui mostra os dados sendo carregados na gui
Return

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
   ; global PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
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
   ; msgbox % PlanilhaLink PlanilhaQuery PlanilhaTipoExportacao PlanilhaRange PlanilhaNomeId
   PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   ; msgbox %PlanilhaLink%
   If(PlanilhaLink)
   {
      ; GoSub, AtualizarPlanilha
      GS_GetCSV_ToListView(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)

      global posicaoColunaNome := GS_GetCSV_Column(, ".*Nome.*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).ColumnPosition
      global posicaoColunaURL := GS_GetCSV_Column(, "i).*(URL|Site|link).*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).ColumnPosition
      ; global planilha := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
      ; Run %iniFile%
   }
Return
/*
   * ESCREVER NO ARQUIVO DE CONFIGURAÇÃO
*/
SaveToIniFile:
   Gui Submit
   If(!FileExist(iniPath))
   {
      FileCreateDir, %appdata% ; criar a pasta
      FileAppend, "" ,iniPath ; criar o arquivo caso ñ exista
   }
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
   GoSub, ReadIniFile
   Notify().AddWindow("Configuração atualizada!`nClique no botão Atualizar para atualizar os dados!",{Time:5000,Icon:177,Background:"0x039018",Title:"SUCESSO",TitleColor:"0xF0F8F1", TitleSize:15, Size:15, Color: "0xF0F8F1"},"","setPosBR")
; global posicaoColunaNome := GS_GetCSV_Column(, ".*Nome.*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).ColumnPosition
; global posicaoColunaURL := GS_GetCSV_Column(, "i).*(URL|Site|link).*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId).ColumnPosition
; global planilha := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
; checkSpreadsheetLink(PlanilhaLink)
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
      ; msgbox %PlanilhaTipoExportacao%
      ; msgbox % PlanilhaLink PlanilhaQuery PlanilhaTipoExportacao PlanilhaRange PlanilhaNomeId
      ; msgbox % PlanilhaLink
      ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
      /*
         * capturar o nome da planilha pela gui/arquivo de configuração .ini
         * se o valor "abaPlanilha" estiver vazio no arquivo de configuração, capturar o nome da planilha pela URL da planilha.
      */
   ; msgbox %PlanilhaNomeId%
   ; msgbox %capture_sheetURL_name1%
   RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   ; msgbox %capture_sheetURL_key1%
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   If(PlanilhaNomeId)
      capture_sheetURL_name := PlanilhaNomeId
   Else
      capture_sheetURL_name := capture_sheetURL_name1
   ; msgbox % capture_sheetURL_name
   ; msgbox % capture_sheetURL_name
   ; msgbox % capture_sheetURL_key1
   ; msgbox % capture_sheetURL_name1
   ; fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key1 "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&sheet=" capture_sheetURL_name "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key1 "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&gid=" capture_sheetURL_name "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)
   ; msgbox % fullSheetURL
   ; CLIPboard := fullSheetURL

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
   ; msgbox % PlanilhaLink PlanilhaQuery PlanilhaTipoExportacao PlanilhaRange PlanilhaNomeId
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   ; msgbox % capture_sheetURL_key1
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   ; msgbox % capture_sheetURL_name1
   ; fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key1 "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&sheet=" capture_sheetURL_name1 "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)
   ; msgbox %PlanilhaTipoExportacao% %PlanilhaLink% %PlanilhaNomeId% %PlanilhaRange% %PlanilhaQuery%

   sheetData_All := GS_GetCSV(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, capture_sheetURL_name1)

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
       LV_ModifyCol(2, "left")
       LV_ModifyCol(2)
       LV_ModifyCol(2, "left")
       LV_ModifyCol(3, "200 Left")
       LV_ModifyCol(4, "70 Left")

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
GS_GetListView_Click(idioma, PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId, regexFindColumnName:= ".*Nome.*", regexFindColumnURL := "i).*(URL|Link).*", action := "openLink", listViewEnterKey := ""){
   Gui Submit, NoHide
   ; msgbox % listViewEnterKey ; apertoUeNTER NA LISTVIEW
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   ; * CAPTURAR A LINHA SELECIONADA NA LISTVIEW
   NumeroLinhaSelecionada := LV_GetNext()
   ; msgbox % NumeroLinhaSelecionada
   ; * Pesquisar por coluna específica
   getColumnName := GS_GetCSV_Column(, regexFindColumnName,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   getColumnURL := GS_GetCSV_Column(, regexFindColumnURL, PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   ; msgbox % getColumnName.arrColumn[2]

   posicaoColunaNome := getColumnName.ColumnPosition
   posicaoColunaURL := getColumnURL.ColumnPosition
   ; msgbox % posicaoColunaURL
   valueColunaNome := getColumnName.ColumnName
   valueColunaURL := getColumnURL.ColumnName
   ; * CAPTURAR VALOR DA COLUNA "NOME"
   LV_GetText(TextoLVNome, NumeroLinhaSelecionada, posicaoColunaNome)
   ; * CAPTURAR VALOR DA COLUNA "URL"
   LV_GetText(TextoLVURL, NumeroLinhaSelecionada, posicaoColunaURL)

   ; * SOLUÇÃO PARA NÃO DEPENDER DA COLUNA URL QUE ESTÁ NA GUI, PEGAR DIRETO DA PLANILHA(array que foi salvo)
   If(!TextoLVURL)
      TextoLVURL := getColumnURL.arrColumnSanitize[NumeroLinhaSelecionada+1]
   ; msgbox % TextoLVURL
   ; msgbox % getColumnURL.arrColumnSanitize[NumeroLinhaSelecionada+1]

   ; msgbox % A_GuiEvent
   if(A_GuiEvent = "DoubleClick" && action = "openLink" || listViewEnterKey = "apertouEnter"){ ; abrir link normal
      /*
         * ABRIR OS LINKS/URLS/DOCUMENTAÇÕES NO NAVEGADOR
         ! IMPORTANTE: Caso tenha mais de um link na coluna, transformar em um array e fazer um loop para abrir os links
      */
      URLSanitized := StrReplace(TextoLVURL, "%idiomapt%", idioma)
      ; msgbox % URLSanitized
      For Index, URL in StrSplit(URLSanitized, " | ")
         {
            If(InStr(URL, "https://www.notion.so/"))
            {
               URLNotion := StrReplace(URL, "https://", "notion://")
               ; getColumnName := GS_GetCSV_Column(, "i).*(notion|anotacoes|anotacao|notes|note|anotações|anotação).*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
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
               Run %ComSpec% /c C:\Users\felipe\AppData\Local\Programs\Notion\Notion.exe "%URLNotion%", , Hide
               RunAs
               WinActivate, Notion
            }
            Else
              Run, %URL%
         }
         Return
   }else if(A_GuiEvent = "R"){ ; CLIQUE COM BOTÃO DIREITO DO MOUSE
   ; * Pesquisar por coluna específica
   getColumnCode := GS_GetCSV_Column(, "i).*(code|codigo|código|source-code|source).*", PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   ; msgbox % getColumnName.arrColumn[2]

   posicaoColunaCode := getColumnCode.ColumnPosition
   valueColunaCode := getColumnCode.ColumnName
    ; URLCode := getColumnCode.arrColumnSanitize[NumeroLinhaSelecionada+1]
    ; * CAPTURAR VALOR DA COLUNA "URL"
    LV_GetText(URLCode, NumeroLinhaSelecionada, posicaoColunaCode)

    ; * SOLUÇÃO PARA NÃO DEPENDER DA COLUNA URL QUE ESTÁ NA GUI, PEGAR DIRETO DA PLANILHA(array que foi salvo)
    If(!URLCode)
      URLCode := getColumnCode.arrColumnSanitize[NumeroLinhaSelecionada+1]
   ; msgbox % TextoLVURL
   ; msgbox % getColumnURL.arrColumnSanitize[NumeroLinhaSelecionada+1]
   ; msgbox %URLCode%
   If(URLCode)
   {
      ; UrlDownloadToFile, %URLCode%, arquivo.txt
      whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
      whr.Open("GET", URLCode, true)
      whr.Send()
      ; Using 'true' above and the call below allows the script to remain responsive.
      whr.WaitForResponse()
      code := whr.ResponseText
      ; MsgBox % code
      Clipboard := code
      ; * Abrir URL de edição do GIST profile default do chrome
      gistEditUrl :=  RegExReplace(URLCode, "/raw/.*", "/edit") ; abrir modo de edição do raw
      Run, "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" "%gistEditUrl%"

      MsgBox, 4160 , SUCESSO!, Código copiado para a área de transferência, 2
      ; * EXIBIR CODIGO NA TELA, FUNÇÕES ESTÃO NO FINAL DO ARQUIVO
      ; displayNum := 0
      ; visibleState := true
      pasteToScreen()
   }Else{
      Notify().AddWindow("Não existe nenhum código para o campo selecionado!!",{Time:2000,Icon:177,Background:"0x039018",Title:"INFO!",TitleSize:15, Size:14, Color: "0xFFFF", TitleColor: "0xE1B9A4"},,"setPosBR")

   }
      ; /*
      ;    ABRIR NOTION
      ; */
      ; ; getColumnName := GS_GetCSV_Column(, "i).*(notion|anotacoes|anotacao|notes|note|anotações|anotação).*",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
      ; if(A_UserName == "Felipe" || A_UserName == "estudos" || A_UserName == "Estudos")
      ;    {
      ;      user := A_UserName
      ;      pass := "xrlo1010"
      ;    }
      ;  Else
      ;    {
      ;      user := "felipe.lullio@hotmail.com"
      ;      pass := "XrLO1000@1010"
      ;    }
      ; RunAs, %user%, %pass%
      ; ; Run, C:\Users\felipe\AppData\Local\Programs\Notion\Notion.exe
      ; Run %ComSpec% /c C:\Users\felipe\AppData\Local\Programs\Notion\Notion.exe "%TextoLinhaSelecionadaNotion%", , Hide
      ; RunAs
      ; WinActivate, Notion
   }
}

      /*
         * FUNÇÃO PARA CRIAR TODAS AS TABS E LISTAS COMBOBOX DE FORMA DINÂMICA
         * ESTOU CAPTURANDO A POSIÇÃO DA COLUNA "URL" E POSIÇÃO DA COLUNA "NOME", LÁ EM CIMA, COMO VARIÁVEL GLOBAL
      */


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
      ; if RegExMatch(y, "im).*" VarPesquisarDados ".*") ; se encontrar o texto digitado no searchbox na linha
      if (RegExMatch(y, "im)" VarPesquisarDados) && x>1) ; x>1 para nao pegar o header (?<!https:\/\/www\.)notion
         {
         row := [], ++cnt
         loop, parse, y, CSV ; dividir a linha em células
               ; if (a_index <= 13)																	;or if a_index in 1,4,5
               row.push(a_loopfield)
         LV_add("",row*)
         }
   SB_SetText("Match(es) da última Pesquisa: " cnt,  4)
   ; loop, % lv_getcount("col")
   ; LV_ModifyCol(a_index,"AutoHdr")
   ; LV_ModifyCol(1, "30 right")
   GuiControl, +Redraw, LVAll
   GuiControl, Focus, LVAll ; dar foco na listview após pesquisar
   LV_Modify(1, "+Select") ; selecionar primeiro item da listview
   LV_ModifyCol()
   i++
   If(LV_GetCount() = 0){
      MsgBox, 4112 , Erro!, A Pesquisa não retornou nada`nAtualizando...!, 2
      GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
      ; Sleep, 500
      ; Notify().AddWindow("Erro",{Time:3000,Icon:28,Background:"0x990000",Title:"ERRO",TitleSize:15, Size:15, Color: "0xCDA089", TitleColor: "0xE1B9A4"},"w330 h30","setPosBR")
      GuiControl, Focus, BtnPesquisar ; dar foco no botao
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
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   ; RegExMatch(PlanilhaLink, "\/d\/(.+)\/", capture_sheetURL_key)
   ; ; msgbox % capture_sheetURL_key1
   ; RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   ; ; msgbox % capture_sheetURL_name1
   ; fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key1 "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&sheet=" capture_sheetURL_name1 "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)
   ; msgbox %PlanilhaTipoExportacao% %PlanilhaLink% %PlanilhaNomeId% %PlanilhaRange% %PlanilhaQuery%
   ; sheetData_All := GS_GetCSV(capture_sheetURL_key1, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, capture_sheetURL_name1)
   ; PlanilhaLink := checkSpreadsheetLink(PlanilhaLink)
   ; Gui, ListView, LVAll
   LV_Delete() ; deletar todas as linhas
   ; deletar todas as colunas
   Loop, % LV_GetCount("Column")
      LV_DeleteCol(1)
   ; executar a planilha novamente
   GS_GetCSV_ToListView(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
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
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=0"
      ; TEMPLATE 2
      else if(PlanilhaLink = "Documentações Banco de Dados")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1ZmlzAhTGDPCsAz9yHAQGHEGPFdLDh1sCE6D7ePHNLjM/edit#gid=0"
      else if(PlanilhaLink = "Documentações Programas")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1ttLOdD2Mz8yZrsLS5vGHW3ojnkeRUOd1YwhwQ5EGIRY/edit#gid=0"
      ; TEMPLATE 2
      else if(PlanilhaLink = "Documentações Programação")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1TkfWTjHWunj6A13X_cMydXX_UEant4sgMKfqr13mjiU/edit#gid=0"
      else if(PlanilhaLink = "Documentações GAPS")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=218001466"
      else if(PlanilhaLink = "Tudo")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/10HK3v8M6T_qkCGktAvqgH1_nmDRudK2SF20R5UGEgP4/edit#gid=0"
      ; TRATAR PELA URL DA PLANILHA
      Else If(RegExMatch(PlanilhaLink, "i).*docs.google.com/.+\/d\/.+\/"))
         Return linkPlanilha := PlanilhaLink
      Else If(!InStr(PlanilhaLink, "https://docs.google.com/spreadsheets") || RegexMatch(PlanilhaLink, "\s{0,}"))
      {
         ; msgbox hi
         MsgBox, 4112 , Erro na URL do Site!, URL Inválida`n- Copie e Cole uma URL do Google Sheets válida!
         ; Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=1280466043"
         ; GoSub, MenuEditarBase
         ; Resetar/Limpar o valor do ComboBox
         GuiControl,ConfigFile:Choose, PlanilhaLink, 1
         Return
      }
      Return linkPlanilha
}

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
   ; msgbox % PlanilhaLink PlanilhaQuery PlanilhaTipoExportacao PlanilhaRange PlanilhaNomeId
   GS_GetListView_Update(PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   LV_ModifyCol(2, "left")
   LV_ModifyCol(2)
   LV_ModifyCol(2, "left")
Return
      /*
         * CASO O CHECKBOX DE "PESQUISAR POR COLUNA" ESTEJA MARCADO
      */
PesquisarDados:
   Gui Submit, NoHide
      /*
         * HACK / TÉCNICA PARA USAR SOMENTE 1 BOTÃO(PESQUISAR) para fazer as pesquisas e para abrir a documentação caso o usuário tenha apertado enter na listview, ou seja,
         * Quando aperta Enter na gui, ela vai executar o botão que tá com "Default", ou seja, vai executar o botão Pesquisar, então, to na listview, apertei enter, vai checar aqui embaixo, se o foco estiver na listview, abra a documentação, se não, foi executado uma pesquisa mesmo
      */
   GuiControlGet,FocusControl,Focus
   ; msgbox %FocusControl%
   If (FocusControl = "SysListView321")
   {
      if(CheckIdiomaPt)
         GS_GetListView_Click("?hl=pt-br",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId,,,, "apertouEnter")
      Else
         GS_GetListView_Click("?hl=en",PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId,,,, "apertouEnter")
      ;  GuiControl, Focus, BtnPesquisar ; dar foco no botao
   }
   Else If(CheckPesquisarColuna = true) ; se o checkbox estiver marcado
      GS_SearchColumns(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
   Else
      GS_SearchRows(VarPesquisarDados,PlanilhaLink, PlanilhaQuery, PlanilhaTipoExportacao, PlanilhaRange, PlanilhaNomeId)
Return


