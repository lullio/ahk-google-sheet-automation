#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
/*
DICAS:
1. Não pode usar vírgula nos campos da planilha pois vai ser entendido como uma coluna em vez de linha
2. Caso queira inserir mais de um item no campo da planilha, não use vírgula ou quebra de linha para separar, use " | "
3. Os dados CSV retornam entre aspas "", isso é bom para você transformar em arrays e usar como variável javascript

*/
; dropdown 1 - principais cursos
Gui Add, Text,section y+10 , Documentações
Gui, Add, ComboBox, Multi x10 y+10 w510 vCurso hwndCursosIDMain sort, 

/*
COLUNA 1
*/
; dropdown 2 - web dev cursos
Gui Add, Text, section x10, Web Developer
Gui, Add, ComboBox, Multi vCursoWebDev hwndCursosIDDev w250 sort, 
; dropdown 3 - Cursos Analytics
Gui Add, Text,, Analytics / Marketing
Gui, Add, ComboBox, Multi vCursoMkt w250 hwndCursosIDAll sort, 
; dropdown 7 - backend
Gui Add, Text,, Backend / Web Server
Gui, Add, ComboBox, Multi vCursoWebServer hwndCursosIDOutros w250 sort, 
/*
COLUNA 2
*/
; dropdown 4 - javascript cursos
Gui Add, Text, ys, JavaScript All
Gui, Add, ComboBox, Multi w250 vCursoJavaScript hwndCursosIDMkt sort, 
; dropdown 5 - sql banco de dados cursos
Gui Add, Text,, SQL
Gui, Add, ComboBox, Multi vCursoSQL hwndCursosIDOutros w250 sort, 
; dropdown 6 - linux cursos
Gui Add, Text, , Linux Courses
Gui, Add, ComboBox , Multi w250 vCursoLinux hwndCursosIDMkt sort, 

/*
COLUNA 3
*/


; gui, font, S7 ;Change font size to 12
; 2º dropdown js courses
Gui, Add, GroupBox, y+15 xs cBlack r13 w560, Lista dos Cursos
Gui Add, Text, yp+25 xp+11 center, Cursos em Andamento
Gui Add, Text, x+155 center, Cursos do Youtube
Gui Font, S10

Gui Add, ComboBox, Multi xs+10 yp+20 w280 center vCursoAndamento hwndDimensoesID sort,
; Gui Add, ComboBox, Multi xs+10 yp+20 w372 center vCursoAll hwndDimensoesID ,
Gui Add, ComboBox, Multi x+10 w237 center vCursoYoutube hwndDimensoesID sort,
Gui Font,
Gui Add, ListView, vCursoDaLista gListViewListener w530 r10 xs+10 y+10 -readonly grid ,
; LV_Modify()
Gui Font, S6.5
Gui Add, Link, w120 y+3 xp+200 vTotalLinhas center,
Gui Add, Button, x+50 w135 h26 gGS_GetListView_Update, Atualizar Lista

; CARREGAR OS DADOS DOS CURSOS DA PLANILHA ANTES DE EXIBIR A GUI, NÃO VAI TER DELAY
; GoSub, getData

; Botões
gui, font, S11
gui, Add, Button, y+25 xs+15 w250 h35  Default, &Abrir Curso
gui, Add, Button, w150 h35 x+10 , &Abrir Anotações
gui, Add, Button, w95 h35 x+10  Cancel, &Cancelar

; EXIBIR E ATIVAR GUI
GuiControl,Focus,Curso
Gui, Show,, Abrir Curso e Controlar Video - Felipe Lulio

; GoSub, controlVideos
; Ignorar o erro que o ahk dá e continuar executando o script

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

GS_GetCSV_ToListView()
test()
; For index, var in GS_GetCSV_Column(, ".*Nome.*").arrColumn
;    msgbox %var%
Return

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
GS_GetCSV(sheetURL_SQLQuery:="", sheetURL_key:="1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34", sheetURL_name:="", sheetURL_format:="csv", sheetURL_range:=""){
   
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" sheetURL_key "/gviz/tq?tqx=out:" sheetURL_format "&range=" sheetURL_range "&sheet=" sheetURL_name "&tq=" GS_EncodeDecodeURI(sheetURL_SQLQuery)
   ; msgbox % fullSheetURL

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
GS_GetCSV_Column(JS_VariableName:="arr", regexFindColumn := "i).*"){
    Gui Submit, NoHide
    sheetData_All := GS_GetCSV() ; Select * limit 1
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
   VARIÁVEIS QUE CONTÉM OS VALORES DAS COLUNAS DA PRIMEIRA LINHA
*/
global ColumnCategory := GS_GetCSV_Column(, "i)Categoria").arrColumnSanitize ; ColumnData.variavelJavascript ColumnData.arrColumn ColumnData.strColumn
global UniqueColumnCategory := RmvDuplic(ColumnCategory)
; Msgbox % ColumnCategory.arrColumnSanitize[5]

/*
   * FUNÇÃO PARA EXIBIR OS DADOS NA LISTVIEW
*/
GS_GetCSV_ToListView(){
   Gui Submit, NoHide   
    sheetData_All := GS_GetCSV() ; Select * limit 1

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
                URLDocTratada := RegExReplace(NomeDocumentacao, "%idiomapt%", "")
               ;  msgbox % URLDocTratada
            ;  if(NomeDocumentacao != "URL")
            ;     Run % URLDocTratada
          }          
       } ; FIM DO LOOP DA LINHA

       LV_ModifyCol() 
       LV_ModifyCol(3, 50) 
       ; total de linhas
       
       TotalLinhas:
         totalLines := LV_GetCount()
         GuiControl, , TotalLinhas, Total de Linhas: %totalLines%
       Return {nomesColunas: coco, colunasHeader: [ColunaHeader1, ColunaHeader2, ColunaHeader3, ColunaHeader4, ColunaHeader5, ColunaHeader6, ColunaHeader7, ColunaHeader8, ColunaHeader9, ColunaHeader10, ColunaHeader11, ColunaHeader12, ColunaHeader13]}
}
; GS_GetCSV_ToListView()

/*
   * FUNÇÃO PARA CAPTURAR AÇÃO AO CLICAR NA LISTVIEW
*/
GS_GetListView_Click(regexFindColumnName:= ".*Nome.*", regexFindColumnURL := "i).*URL|Link.*", action := "openLink"){
   Gui Submit, NoHide

   ; * CAPTURAR A LINHA SELECIONADA NA LISTVIEW
   NumeroLinhaSelecionada := LV_GetNext() 
   ; * Pesquisar por coluna específica
   getColumnName := GS_GetCSV_Column(, regexFindColumnName)
   getColumnURL := GS_GetCSV_Column(, regexFindColumnURL)

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
      For Index, URL in StrSplit(TextoLVURL, " | ")
         {
              URL := RegExReplace(URL, "%idiomapt%", "")
            ;   msgbox %URL%
              Run, %URL%
         } 
   }else if(A_GuiEvent == "DoubleClick" && action = "openAHKChrome"){ ; abrir ahk chrome
      URL := RegExReplace(TextoLVURL, "%idiomapt%", "")
      if !(PageInst := Chrome.GetPageByURL(URL, "contains"))
         {
            ChromeInst := new Chrome(profileName,URL,"--remote-debugging-port=9222 --remote-allow-origins=* --profile-directory=""Default""",chPath)
            Notify().AddWindow("Não encontrei o site aberto no Chrome, Vou abrir pra você agora!",{Time:6000,Icon:28,Background:"0x900C3F",Title:"OPS!",TitleSize:15, Size:15, Color: "0xCDA089", TitleColor: "0xE1B9A4"},,"setPosBR")
            Sleep, 500
            contador1 := 0
            while !(PageInst)
            {
               Sleep, 500
               Notify().AddWindow("procurando instância do chrome...!",{Time:6000,Icon:28,Background:"0x1100AA",Title:"ERRO!",TitleSize:15, Size:15, Color: "0xCDA089", TitleColor: "0xE1B9A4"},,"setPosBR")
               PageInst := Chrome.GetPageByURL(URL, "contains")
               contador1++
               if(contador1 >= 30){
                  PageInst.Disconnect()
                  break
               }
            }
         }
         Sleep, 500
         ; aqui está o fix pra esperar a página carregar
         PageInst := Chrome.GetPageByURL(URL, "contains")
         Sleep, 500
      ; SUPER IMPORTANTE, ATIVAR A TAB/PÁGINA, ACTIVATE, FOCUS
         PageInst.Call("Page.bringToFront")
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


test(){
global ColumnCategory := GS_GetCSV_Column(, "i)Categoria").arrColumnSanitize ; ColumnData.variavelJavascript ColumnData.arrColumn ColumnData.strColumn
global UniqueColumnCategory := RmvDuplic(ColumnCategory)

sheetData_All := GS_GetCSV() ; Select * limit 1
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
     if(RegExMatch(cellContent, category)) ; se for a 1ª linha header e texto for igual a "nome"
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
      getColumnName := GS_GetCSV_Column(, regexFindColumnName)
      getColumnURL := GS_GetCSV_Column(, regexFindColumnURL)

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

AbrirCurso:
Gui, Submit, NoHide
AHK_GetControls()
Return

/*
   ----
   ----
   * LABELS
*/
; LABEL PARA CAPTURAR CLIQUE NA LISTVIEW
ListViewListener:
   GS_GetListView_Click()
Return
; LABEL PARA CAPTURAR O CLIQUE NO BOTÃO ATUALIZAR LISTA
GS_GetListView_Update:
      LV_Delete()
      GS_GetCSV_ToListView()
Return

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

;  getData:
;     Gui Submit, NoHide
;           ; query para selecionar apenas a primeira coluna
;           urlDataQueryGA4 := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/gviz/tq?tqx=out:csv&range=A:D&sheet=All-Docs&tq=select%20*%20where%20D%20matches%20'%5EGA4.*'%20AND%20D%20is%20not%20null"
;           urlDataQueryGA4 := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/gviz/tq?tqx=out:csv&range=A:D&sheet=All-Docs&tq=select%20*%20where%20D%20matches%20'%5EGA4.*'%20AND%20D%20is%20not%20null"
;          ;  urlDataQueryGA4 := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/gviz/tq?tqx=out:csv&range=A:D&sheet=All-Docs&tq=select%20*%20where%20B%20contains%20'Insight'"

 
;     ;! CAPTURAR TODAS LINHAS E COLUNAS DA PLANILHA
;     dataAllRows := GS_GetCSV(urlDataQueryGA4)
;     msgbox % dataAllRows
;     rowsNome := []
;     teste := ""
;     Loop, parse, dataAllRows, `n ; PROCESSAR CADA LINHA DA TABELA/PLANILHA
;        {
;           LineNumber := A_Index ; Index da linha
;           LineContent := A_LoopField ; Conteúdo da linha, todos valores da linha, a 1ª linha vai ser o HEADER(vc consegue capturar os headers das colunas)
          
;        Loop, parse, A_LoopField, `, ; PROCESSAR CADA CÉLULA/CAMPO DA LINHA ATUAL
;        {
;          ColumnNumber := A_Index ; Index da linha
;          cellContent := A_LoopField ; armazenar o conteúdo da célula numa variável
;          ; msgbox %A_LoopField% ; Exibe cada célula, cada camnpo da planilha
;          ; msgbox % SubStr(A_LoopField, 2,-1) ; remove o primeiro e último catactere (as aspas)
;          if(InStr(cellContent, "Nome")) ; se for a 1ª linha header e texto for igual a "nome"
;          {
            
;             Loop, parse, dataAllRows, `n
;                {
;                /*
;                   SALVAR TODAS AS LINHAS DA COLUNA "Nome"
;                */
;                ; msgbox %A_LoopField% ; aqui exibe a linha inteira
;                ; msgbox % columnData := StrSplit(A_LoopField,",")[ColumnNumber] ; Somente o valor da celula da coluna
;                rowsNome.push(StrSplit(A_LoopField,",")[ColumnNumber]) ; Somente o valor da celula da coluna
;                test.= StrSplit(A_LoopField,",")[ColumnNumber] ", "
;                ; ifequal,a_index,13,break ;prevents from reading columns that are further out se chegar na linha 13 quebrar
;                ; ColunaNome := RegExReplace([1], aspa , "") ; 1ª coluna da planilha
;                }
;             ; msgbox "coluna nome " %A_LoopField%
;          }
;          ; msgbox % StrSplit(LineContent,",")[LineNumber]
;        }
       
;       ;  If(StrSplit(A_LoopField,",")[A_Index]){
;       ;  }

;          /*
;             COLUNAS DA PLANILHA
;          */
;           Coluna1 := RegExReplace(StrSplit(A_LoopField,",")[1], aspa , "") ; 1ª coluna da planilha
;          ;  msgbox %Coluna1%
;           Coluna2 := RegExReplace(StrSplit(A_LoopField,",")[2], aspa , "") ; 2ª coluna da planilha
;           Coluna3 := RegExReplace(StrSplit(A_LoopField,",")[3], aspa , "") ; 3ª coluna da planilha
;           Coluna4 := RegExReplace(StrSplit(A_LoopField,",")[4], aspa , "") ; 4ª coluna da planilha
;           Coluna5 := RegExReplace(StrSplit(A_LoopField,",")[5], aspa , "") ; 5ª coluna da planilha
;           Coluna6 := RegExReplace(StrSplit(A_LoopField,",")[6], aspa , "") ; 6ª coluna da planilha
;           Coluna7 := RegExReplace(StrSplit(A_LoopField,",")[7], aspa , "") ; 7ª coluna da planilha
;          /*
;             ADICIONAR AS LINHAS/COLUNAS NA PRIMEIRA LISTIVEW DA GUI
;          */
;           ; LV_Add("" , Coluna1, SubStr(Coluna2, 2,-1), SubStr(Coluna3, 2,-1), SubStr(Coluna4, 2,-1), SubStr(Coluna5, 2,-1)) ; serve para remover as aspas na frente e final         
;           LV_Add("" , Coluna1, Coluna2, Coluna3, Coluna4)
          
;          /*
;             O CONTEÚDO NA PLANILHA POSSUI OS TEXTOS "%idiomapt%", vamos tratar isso para não ser considerado um erro na url
;          */
;           For Index, NomeDocumentacao in StrSplit(Coluna3, " | ")
;           {
;                ;  msgbox % index " is " NomeDocumentacao 
;                 URLDocTratada := RegExReplace(NomeDocumentacao, "%idiomapt%", "")
;                ;  msgbox % URLDocTratada
;             ;  if(NomeDocumentacao != "URL")
;             ;     Run % URLDocTratada
;           }
;           /*
;             !ORGANIZAR O CONTEÚDO DA PLANILHA POR CATEGORIA, SEPARAR CADA NOME NA SUA DEVIDA COMBOBOX/LISTBOX
;             msgbox % StrSplit(Coluna3, " | ").MaxIndex() ; exibir o tamanho do array
;           */
;           ListAllCourses .= RegexReplace(StrSplit(A_LoopField,",")[1] "|", aspa, "") ; salvar todos os cursos
;           If InStr(Coluna3, "sql")
;              ListSQLCourses .= RegexReplace(StrSplit(A_LoopField,",")[1] "|", aspa, "")
;           If InStr(Coluna3, "web-dev")
;              ListWebDevCourses .= RegexReplace(StrSplit(A_LoopField,",")[1] "|", aspa, "")
;           If InStr(Coluna3, "javascript") || If InStr(Coluna3, "js-frameworks") 
;              ListJavaScriptCourses .= RegexReplace(StrSplit(A_LoopField, ",")[1] "|", aspa, "")
;           If InStr(Coluna3, "analytics") || InStr(Coluna3, "ads") || InStr(Coluna3, "wordpress") 
;              ListAnalyticsCourses .= RegexReplace(StrSplit(A_LoopField, ",")[1] "|", aspa, "")
;           If InStr(Coluna3, "linux") || InStr(Coluna3, "redes") || InStr(Coluna3, "hacking") 
;              ListLinuxCourses .= RegExReplace(StrSplit(A_LoopField, ",")[1] "|", aspa, "")
;           If InStr(Coluna3, "top-rated") 
;              ListTopCourses .= RegexReplace(StrSplit(A_LoopField, ",")[1] "|", aspa, "")
;           If InStr(Coluna3, "web-server") 
;              ListWebServerCourses .= RegexReplace(StrSplit(A_LoopField, ",")[1] "|", aspa, "")
;           If InStr(Coluna3, "em-andamento") 
;              ListAndamentoCourses .= RegexReplace(StrSplit(A_LoopField, ",")[1] "|", aspa, "")
;           If InStr(Coluna2, "youtube.com") 
;              ListYoutubeCourses .= RegexReplace(StrSplit(A_LoopField, ",")[1] "|", aspa, "")
 
;        } 
;        msgbox % rowsNome[2]
;        msgbox % "let arr = " RegExReplace(test, ",\s+$", ";")
;        Clipboard := "let arr = " RegExReplace(test, " ?,\s+$", ";")
;        ; MODIFICANDO TODAS COMBOBOX PARA POPULAREM OS DADOS DA PLANILHA
;        GuiControl,1:, Curso, %ListTopCourses% ; main courses
;        GuiControl,1:, CursoWebDev, %ListWebDevCourses% ; web dev courses
;        GuiControl,1:, CursoJavaScript, %ListJavaScriptCourses% ; analytics mkt courses
;        GuiControl,1:, CursoMkt, %ListAnalyticsCourses% ; analytics mkt courses
;        GuiControl,1:, CursoSQL, %ListSQLCourses% ; analytics mkt courses
;        GuiControl,1:, CursoWebServer, %ListWebServerCourses% ; analytics mkt courses
;        GuiControl,1:, CursoLinux, %ListLinuxCourses% ; analytics mkt courses
;        GuiControl,1:, CursoAll, %ListAllCourses% ; analytics mkt courses
;        GuiControl,1:, CursoAndamento, %ListAndamentoCourses% ; cursos em andamento
;        GuiControl,1:, CursoYoutube, %ListYoutubeCourses% ; cursos do youtube
 
;        LV_ModifyCol()
 

;        ; exibir total de linhas
;        TotalLinhas2:
;           totalLines := LV_GetCount()
;           GuiControl, , TotalLinhas, Total de Cursos: %totalLines%
;        Return
;  Return