#Include <Default_Settings>
SetWorkingDir, %A_ScriptDir%

Menu, Tray, Icon, C:\Windows\system32\shell32.dll,68 ;Set 

/*
; https://www.autohotkey.com/boards/viewtopic.php?t=76770#:~:text=Hotkeys%3A,text)%20in%20Clipboard%20to%20screen.

- F3 = Colar imagem do clipboard

- Mouse Scroll = mudar tamanho do paste, middle button resetar o tamanho (só funciona para imagems)

- Ctrl+mouse scroll = mudar a transparência | Ctrl+ middlebutton resetar a transparência

- Drag the paste to move it
- Double click the paste to close it

- Ctrl+Alt+3 = paste unlickable(you can clickthrough the paste)

- Shift+F3 = hide / show all paste
- Ctrl+f3 = close al paste
- Ctrl+shift+f3 = exit script

- F4 = Abrir imagem de arquivo
- CTRL+P = colar imagem especifica de pasta
Adicionar minimizar um paste

ahk
*/
if not A_IsAdmin
   Run *RunAs "%A_ScriptFullPath%" ; (A_AhkPath is usually optional if the script has the .ahk extension.) You would typically check  first.p

displayNum := 0
visibleState := true

F3::
	pasteToScreen(){
		if DllCall("IsClipboardFormatAvailable", "UInt", 1)
			displayText(Clipboard)
		If DllCall("IsClipboardFormatAvailable", "UInt", 2){
			if DllCall("OpenClipboard", "uint", 0) {
				hBitmap := DllCall("GetClipboardData", "uint", 2)
				DllCall("CloseClipboard")
			}
			displayImg(hBitmap)
		}
		if DllCall("IsClipboardFormatAvailable", "UInt", 15){
			imgFile := Clipboard
			if(hBitmap := LoadPicture(imgFile))
				displayImg(hBitmap)
		}
	}

; Alterar e incluir as opções da imagem?
displayText(text){
	global
	Gui, New, +hwndpasteText%displayNum% -Caption +AlwaysOnTop +ToolWindow -DPIScale
	local textHnd := pasteText%displayNum%
	Gui, Margin, 10, 10
	Gui, Font, s16
	Gui, Add, Text,, % text
	OnMessage(0x201, "move_Win")
	OnMessage(0x203, "close_Win")
	Gui, Show,, pasteToScreen_text
	transparency%textHnd% := 100
	displayNum++
}

displayImg(hBitmap){
	global
	Gui, New, +hwndpasteImg%displayNum% -Caption +AlwaysOnTop +ToolWindow -DPIScale
	local imgHnd := pasteImg%displayNum%
	Gui, Margin, 0, 0
	Gui, Add, Picture, Hwndimg%imgHnd%, % "HBITMAP:*" hBitmap
	OnMessage(0x201, "move_Win")
	OnMessage(0x203, "close_Win")
	Gui, Show,, pasteToScreen_img
	local img := img%imgHnd%
	ControlGetPos,,, width%imgHnd%, height%imgHnd%,, ahk_id %img%
	scale%imgHnd% := 100
	transparency%imgHnd% := 100
	displayNum++
}

move_Win(){
	PostMessage, 0xA1, 2
}

close_Win(){
	id := WinExist("A")
	transparency%id% := ""
	scale%id% := ""
	width%id% := ""
	height%id% := ""
	Gui, Destroy
}

#IfWinActive pasteToScreen

^WheelDown::
	decreaseTransparency(){
		id := WinExist("A")
		transparency%id% -= 5
		If (transparency%id% < 10)
			transparency%id% = 10
		transparency := transparency%id% * 255 // 100
		WinSet, Transparent, %transparency%, A
		tooltip, % "Opacity:" transparency%id% "%"
		SetTimer, RemoveToolTip, -500
	}

^WheelUp::
	increaseTransparency(){
		id := WinExist("A")
		transparency%id% += 5
		If (transparency%id% > 100)
			transparency%id% = 100
		transparency := transparency%id% * 255 // 100
		WinSet, Transparent, %transparency%, A
		tooltip, % "Opacity:" transparency%id% "%"
		SetTimer, RemoveToolTip, -500
	}

^MButton::
	resetTransparency(){
		id := WinExist("A")
		transparency%id% = 100
		WinSet, Transparent, 255, A
		tooltip, % "Opacity:" transparency%id% "%"
		SetTimer, RemoveToolTip, -500
	}

#IfWinActive pasteToScreen_img

~WheelDown::
	decreaseSize(){
		id := WinExist("A")
		img := img%id%
		scale%id% -= 10
		If (scale%id% < 10)
			scale%id% = 10
		WinGetPos,,, width, height
		width := width%id% * scale%id% // 100
		height := height%id% * scale%id% // 100
		GuiControl, MoveDraw, %img%, w%width% h%height%
		WinMove,,,,, width, height
		tooltip, % "Size:" scale%id% "%"
		SetTimer, RemoveToolTip, -500
	}

~WheelUp::
	increaseSize(){
		id := WinExist("A")
		img := img%id%
		scale%id% += 10
		WinGetPos,,, width, height
		width := width%id% * scale%id% // 100
		height := height%id% * scale%id% // 100
		GuiControl, MoveDraw, %img%, w%width% h%height%
		WinMove,,,,, width, height
		tooltip, % "Size:" scale%id% "%"
		SetTimer, RemoveToolTip, -500
	}

~MButton::
	resetSize(){
		id := WinExist("A")
		img := img%id%
		scale%id% = 100
		width := width%id%
		height := height%id%
		GuiControl, MoveDraw, %img%, w%width% h%height%
		WinMove,,,,, width, height
		tooltip, % "Size:" scale%id% "%"
		SetTimer, RemoveToolTip, -500
	}

#IfWinActive

^!3::
	toggleClickThroughState(){
		WinGet, id, List, pasteToScreen
		Loop, %id%
		{
			this_id := id%A_Index%
			WinSet, ExStyle, ^0x20, ahk_id %this_id%
		}
	}

+F3::
	toggleVisibleState(){
		global visibleState
		if(visibleState){
			WinGet, id, List, pasteToScreen
			Loop, %id%
			{
				this_id := id%A_Index%
				WinHide, ahk_id %this_id%
			}
			visibleState := false
		} else {
			DetectHiddenWindows, On
			WinGet, id, List, pasteToScreen
			Loop, %id%
			{
				this_id := id%A_Index%
				WinShow, ahk_id %this_id%
			}
			DetectHiddenWindows, Off
			visibleState := true
		}
	}



; Colar imagem específica
; ^p::
; hBitmap := LoadPicture("C:\Users\Estudos\Downloads\huge.png")
; Notify().addWindow("", {Time: 2000, Flash:1000, FlashColor: "0x1100AA",Icon:300, IconSize: 40, Background:"0x1100AA",Title:"Imagem carregada",TitleSize:16,size:12, Sound:800,Hide:"Right" })
; displayImg(hBitmap)
; Return

#F3::
jpegFilePath := A_Desktop . "\MyImageFile.jpeg" ; specify the file path you prefer
quality := 1 ; specify quality from 0 to 1, where 1 is 100%

hBitmap := GetBitmapFromClipboard()
SaveBitmapToJpeg(hBitmap, jpegFilePath, quality)
DllCall("DeleteObject", "Ptr", hBitmap)

GetBitmapFromClipboard() {
   static CF_BITMAP := 2, CF_DIB := 8, SRCCOPY := 0x00CC0020
   if !DllCall("IsClipboardFormatAvailable", "UInt", CF_BITMAP)
      throw "Não encontrei imagem no desktop"
   if !DllCall("OpenClipboard", "Ptr", 0)
      throw "OpenClipboard failed"
   hDIB := DllCall("GetClipboardData", "UInt", CF_DIB, "Ptr")
   hBM := DllCall("GetClipboardData", "UInt", CF_BITMAP, "Ptr")
   DllCall("CloseClipboard")
   if !hDIB
      throw "GetClipboardData failed"
   pDIB := DllCall("GlobalLock", "Ptr", hDIB, "Ptr")
   width := NumGet(pDIB + 4, "UInt")
   height := NumGet(pDIB + 8, "UInt")
   bpp := NumGet(pDIB + 14, "UShort")
   DllCall("GlobalUnlock", "Ptr", pDIB)

   hDC := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
   oBM := DllCall("SelectObject", "Ptr", hDC, "Ptr", hBM, "Ptr")

   hMDC := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
   hNewBM := CreateDIBSection(width, -height,, bpp)
   oPrevBM := DllCall("SelectObject", "Ptr", hMDC, "Ptr", hNewBM, "Ptr")
   DllCall("BitBlt", "Ptr", hMDC, "Int", 0, "Int", 0, "Int", width, "Int", height
      , "Ptr", hDC , "Int", 0, "Int", 0, "UInt", SRCCOPY)
   DllCall("SelectObject", "Ptr", hDC, "Ptr", oBM, "Ptr")
   DllCall("DeleteDC", "Ptr", hDC), DllCall("DeleteObject", "Ptr", hBM)
   DllCall("SelectObject", "Ptr", hMDC, "Ptr", oPrevBM, "Ptr")
   DllCall("DeleteDC", "Ptr", hMDC)
   Return hNewBM
}

CreateDIBSection(w, h, ByRef ppvBits := 0, bpp := 32) {
   hDC := DllCall("GetDC", "Ptr", 0, "Ptr")
   VarSetCapacity(BITMAPINFO, 40, 0)
   NumPut(40 , BITMAPINFO, 0)
   NumPut( w , BITMAPINFO, 4)
   NumPut( h , BITMAPINFO, 8)
   NumPut( 1 , BITMAPINFO, 12)
   NumPut(bpp, BITMAPINFO, 14)
   hBM := DllCall("CreateDIBSection", "Ptr", hDC, "Ptr", &BITMAPINFO, "UInt", 0
      , "PtrP", ppvBits, "Ptr", 0, "UInt", 0, "Ptr")
   DllCall("ReleaseDC", "Ptr", 0, "Ptr", hDC)
   return hBM
}

SaveBitmapToJpeg(hBitmap, destJpegFilePath, quality := 0.75) {
    static CLSID_WICImagingFactory  := "{CACAF262-9370-4615-A13B-9F5539DA4C0A}"
         , IID_IWICImagingFactory   := "{EC5EC8A9-C395-4314-9C77-54D7A935FF70}"
         , GUID_ContainerFormatJpeg := "{19E4A5AA-5662-4FC5-A0C0-1758028E1057}"
         , WICBitmapIgnoreAlpha := 0x2, GENERIC_WRITE := 0x40000000, VT_R4 := 0x00000004
         , WICBitmapEncoderNoCache := 0x00000002, szPROPBAG2 := 24 + A_PtrSize*2

   VarSetCapacity(GUID, 16, 0)
   DllCall("Ole32\CLSIDFromString", "WStr", GUID_ContainerFormatJpeg, "Ptr", &GUID)
   IWICImagingFactory := ComObjCreate(CLSID_WICImagingFactory, IID_IWICImagingFactory)
   Vtable( IWICImagingFactory    , CreateBitmapFromHBITMAP := 21 ).Call("Ptr", hBitmap, "Ptr", 0, "UInt", WICBitmapIgnoreAlpha, "PtrP", IWICBitmap)
   Vtable( IWICImagingFactory    , CreateStream            := 14 ).Call("PtrP", IWICStream)
   Vtable( IWICStream            , InitializeFromFilename  := 15 ).Call("WStr", destJpegFilePath, "UInt", GENERIC_WRITE)
   Vtable( IWICImagingFactory    , CreateEncoder           :=  8 ).Call("Ptr", &GUID, "Ptr", 0, "PtrP", IWICBitmapEncoder)
   Vtable( IWICBitmapEncoder     , Initialize              :=  3 ).Call("Ptr", IWICStream, "UInt", WICBitmapEncoderNoCache)
   Vtable( IWICBitmapEncoder     , CreateNewFrame          := 10 ).Call("PtrP", IWICBitmapFrameEncode, "PtrP", IPropertyBag2)

   Vtable( IPropertyBag2         , CountProperties         :=  5 ).Call("UIntP", count)
   VarSetCapacity(arrPROPBAG2    , szPROPBAG2*count, 0)
   Vtable( IPropertyBag2         , GetPropertyInfo         :=  6 ).Call("UInt", 0, "UInt", count, "Ptr", &arrPROPBAG2, "UIntP", read)
   Loop % read
      addr := &arrPROPBAG2 + szPROPBAG2*(A_Index - 1)
   until StrGet(NumGet(addr + 8 + A_PtrSize)) = "ImageQuality" && found := true
   if found {
      VarSetCapacity(variant, 24, 0)
      NumPut(VT_R4, variant)
      NumPut(quality, variant, 8, "Float")
      Vtable( IPropertyBag2, Write := 4 ).Call("UInt", 1, "Ptr", addr, "Ptr", &variant)
   }
   Vtable( IWICBitmapFrameEncode , Initialize              :=  3 ).Call("Ptr", IPropertyBag2)
   Vtable( IWICBitmapFrameEncode , WriteSource             := 11 ).Call("Ptr", IWICBitmap, "Ptr", 0)
   Vtable( IWICBitmapFrameEncode , Commit                  := 12 ).Call()
   Vtable( IWICBitmapEncoder     , Commit                  := 11 ).Call()
   for k, v in [IWICBitmapFrameEncode, IWICBitmapEncoder, IPropertyBag2, IWICStream, IWICBitmap, IWICImagingFactory]
      ObjRelease(v)
}

Vtable(ptr, n) {
   return Func("DllCall").Bind(NumGet(NumGet(ptr+0), A_PtrSize*n), "Ptr", ptr)
}
Return

^F3::
	destroyAllPaste(){
		WinGet, id, List, pasteToScreen
		Loop, %id%
		{
			this_id := id%A_Index%
			SendMessage, 0x203,,,, ahk_id %this_id%
		}
	}

/* MINHAS MODIFICAÇÕES
*/
; Abrir imagem de arquivo
F6::
	FileSelectFile, imgFile, 3
	hBitmap := LoadPicture(imgFile)
	displayImg(hBitmap)
return

^+F3::ExitApp


; HELP VS CODE
:*:helpvscode::
:*:helpvsc::
:*:atalhosvscode::
:*:img.vscode::
:*:shortcutvscode::
:*:shortcutsvscode::
:*:shortcutsvsc::
:*:vscodeshorcut::
:*:vscodehelp::
:*:vscodeatalho::
#^e::
	hBitmap := LoadPicture("images-paste-to-screen\programs\VSCODE-1.png")
	hBitmap2 := LoadPicture("images-paste-to-screen\programs\VSCODE-2.png")
	displayImg(hBitmap)
	displayImg(hBitmap2)
Return

; HELP CHROME
:*:helpchrome::
:*:helpbrowser::
:*:helpbrowser::
:*:atalhoschrome::
:*:atalhosnavegador::
:*:shortcutnavegador::
:*:shortcutsnavegador::
:*:shortcutschrome::
:*:shortcutchrome::
:*:chromeshortcut::
:*:chromehelp::
:*:chromeatalho::
:*:img.chrome::
#^r::
	hBitmap := LoadPicture("images-paste-to-screen\chrome-browser\chrome-shortcuts.png")
	displayImg(hBitmap)
Return

; HELP WINDOWS
:*:atalhoswindows::
:*:shortcutwindows::
:*:shortcutswindows::
:*:atalhowindows::
:*:helpwindows::
:*:windowsatalho::
:*:windowsshortcut::
:*:windowshelp::
:*:img.windows::
#^w::
	hBitmap := LoadPicture("images-paste-to-screen\windows\WINDOWS-SHORTCUTS1.png")
	displayImg(hBitmap)
Return

; HELP MEUS SCRIPTS
:*:atalhosmeuscript::
:*:shortcutmeuscript::
:*:shortcutsmeuscript::
:*:helpmeuscript::
:*:meuscripthelp::
:*:meuscriptshorcut::
:*:meuscripthelp::
:*:meuscriptatalho::
:*:helpmeucript::
:*:img.scripts::
:*:helpscripts::
#^s::
	hBitmap := LoadPicture("images-paste-to-screen\scripts\MEU-SCRIPT-ATALHOS.png")	
	displayImg(hBitmap)
Return

; HELP MINHAS HOSTRINGS
:*:atalhoshotstring::
:*:shortcuthotstring::
:*:shortcutshotstring::
:*:helphotstring::
:*:hotstringatalho::
:*:hotstringshortcut::
:*:hotstringhelp::
:*:helptextcompl::
:*:img.hotstring::
#^h::
	hBitmap := LoadPicture("images-paste-to-screen\scripts\MEU-SCRIPT-ATALHOS2-HOSTRINGS.png")	
	displayImg(hBitmap)
Return

/*
PROGRAMAÇÃO
*/
; HELP DOM NODES
:*:helpdomnodes::
:*:helpnodes::
:*:helpnode::
:*:nodeshelp::
:*:htmlnodeshelp::
:*:nodesshortcut::
:*:nodesatalho::
:*:htmlnodesshortcut::
:*:htmlnodeshelp::
:*:htmlnodesatalho::
:*:img.nodes::
#^n::
	hBitmap := LoadPicture("images-paste-to-screen\javascript\DOM-NODES1.png")
	hBitmap1 := LoadPicture("images-paste-to-screen\javascript\DOM-NODES2.png")
	displayImg(hBitmap)
	displayImg(hBitmap1)
Return

; HELP HTML5 EMMET
:*:helpemmet::
:*:htmlemmethelp::
:*:helphtmlemmet::
:*:img.helpemmet::
:*:img.emmethtml::
:*:img.htmlemmet::
:*:emmethtml5::
:*:emmethtml::
:*:shortcuthtml::
:*:shortcutshtml::
#^y::
	hBitmap := LoadPicture("images-paste-to-screen\html\EMMET-1.png")
	hBitmap2 := LoadPicture("images-paste-to-screen\html\EMMET-2.png")
	displayImg(hBitmap)
	displayImg(hBitmap2)
Return
	
RemoveToolTip:
	ToolTip
return
