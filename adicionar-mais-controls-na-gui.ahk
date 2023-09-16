#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
; Crie a janela da GUI
Gui, New

; Botão para adicionar controles
Gui, Add, Button, x10 y10 w100 h30 gAdicionarControles, Adicionar Controles

; Mostre a GUI
Gui, Show, w500 h500
return

AdicionarControles:
; Adicione mais campos e controles à GUI após clicar no botão
; Por exemplo, adicionar um campo de texto e um botão
Gui, Add, Text, x10 y50 w120 h100, Novo Campo de Texto:
Gui, Add, Edit, x120 y50 w120 h20 vNovoConteudo
Gui, Add, Button, x10 y80 w100 h30 gMostrarConteudo, Mostrar Conteúdo

; Redesenhe a GUI para exibir os novos controles
Gui, Show, ,w500
return

MostrarConteudo:
; Obtenha o conteúdo do campo de texto e mostre-o em uma caixa de mensagem
GuiControlGet, conteudo, , NovoConteudo
MsgBox, Conteúdo do Campo de Texto: %conteudo%
return