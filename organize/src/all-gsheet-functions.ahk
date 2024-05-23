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
   ; msgbox %capture_sheetURL_key1%
   RegExMatch(PlanilhaLink, "#gid=(.+)", capture_sheetURL_name)
   If(PlanilhaNomeId)
      capture_sheetURL_name := PlanilhaNomeId
   Else
      capture_sheetURL_name := capture_sheetURL_name1
   
   fullSheetURL = % "https://docs.google.com/spreadsheets/d/" capture_sheetURL_key1 "/gviz/tq?tqx=out:" PlanilhaTipoExportacao "&range=" PlanilhaRange "&gid=" capture_sheetURL_name "&tq=" GS_EncodeDecodeURI(PlanilhaQuery)

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
         ;  msgbox % LineNumber
         ;  msgbox % LineContent
         ;  if(InStr(LineContent, "`n"))
            ; msgbox % LineContent

       Loop, parse, A_LoopField, `, ; PROCESSAR CADA CÉLULA/CAMPO DA LINHA ATUAL
       {
         ColumnNumber := A_Index ; Index da coluna
         cellContent := A_LoopField ; armazenar o conteúdo da célula numa variável
         ; msgbox % ColumnNumber
         ;  msgbox %A_LoopField% ; Exibe cada célula, cada camnpo da planilha
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
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1GB5rHO87c-1uGmvF5KTLrRtI1PX2WMdNS93fSdRpy34/edit#gid=0"
      ; TEMPLATE 2
      else if(PlanilhaLink = "Documentações Banco de Dados")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1ZmlzAhTGDPCsAz9yHAQGHEGPFdLDh1sCE6D7ePHNLjM/edit#gid=0"
      else if(PlanilhaLink = "Documentações Programas")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/1ttLOdD2Mz8yZrsLS5vGHW3ojnkeRUOd1YwhwQ5EGIRY/edit#gid=0"
      else if(PlanilhaLink = "Documentações Work")
         Return linkPlanilha := "https://docs.google.com/spreadsheets/d/18cMG-GKYTR7MjKw4NQOGu4El-8n62qncCwyWVIktrTg/edit#gid=0"
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