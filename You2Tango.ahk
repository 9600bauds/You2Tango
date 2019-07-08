#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

IsNum( str ) { ;Fuck AHK.
	if str is number
		return true
	return false
}

PicExists(filename) {
    CoordMode, Pixel, Mouse, Screen
    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %filename%
    if (ErrorLevel = 2){
        MsgBox Could not conduct the search for %filename%.
        return false
    }
    if (ErrorLevel = 0){
        return true
    }
    return false
}

ClickPic(filename, offsetX:=0, offsetY:=0) {  
    CoordMode, Pixel, Mouse, Screen
    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %filename%
    if (ErrorLevel = 2){
        MsgBox Could not conduct the search for %filename%.
        return false
    }
    if (ErrorLevel = 1){
        MsgBox Icon %filename% could not be found on the screen.
        return false
    }
    
    FoundX += offsetX
    FoundY += offsetY
    Click, %FoundX%, %FoundY%
    return true
}


CopiarUnidadMedidaVentas(){
    if(not WinExist("ACTUALIZACION DE ARTICULOS")){
        MsgBox No existe ACTUALIZACION DE ARTICULOS.
        return
    }
    
    WinMenuSelectItem, ACTUALIZACION DE ARTICULOS, , Modificar ;Modo Modificar.
    WinWait, ACTUALIZACION DE ARTICULOS
    
    ;if(not ClickPic("Images/UnidaddeMedidaVentas.png", 175, 5)){return}
    ControlFocus, TEdit4, ACTUALIZACION DE ARTICULOS ;TEdit4 es el ID del campo de texto Unidad de Medida Ventas.    
    SendMessage, 0x301, , , TEdit4, ACTUALIZACION DE ARTICULOS ;"SendMessage, 0x301" envia CTRL+C.
    WinWait, ACTUALIZACION DE ARTICULOS
    
    ControlSend,,{Esc}, ACTUALIZACION DE ARTICULOS ;Cancelar modo Modificar
    WinWait, ACTUALIZACION DE ARTICULOS
}

PegarPrecio98o99(mult:=1){
    Clipboard := RegExReplace(Clipboard, "\R") ;Eliminar líneas extra generadas por Excel y otros programas
    Clipboard := RegExReplace(Clipboard, ",", ".") ;Tango sólo quiere puntos.
    Clipboard := RegExReplace(Clipboard, "\.(?![^.]+$)", "")  ;Quitar todos los puntos excepto el último.
    if(not IsNum(Clipboard)) {
        MsgBox, Clipboard is not a number.
        return
    }
    multiplied := (Clipboard * mult)
    multiplied = % Round(multiplied, 2) ;Tango sólo quiere 2 decimales.

    if(not WinExist("ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO")){
        MsgBox No existe ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO.
        return
    }
    WinActivate, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    
    WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    WinMenuSelectItem, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO, , Modificar
    WinWait, ACTUALIZACION DE ARTICULOS
    if(PicExists("Images/Dolar.png")){
        if(not ClickPic("Images/Dolar.png", 425, 5)){
            return
        }
        WinWait, ACTUALIZACION DE ARTICULOS
        if(not ClickPic("Images/Dolar_Seleccionado.png", 425, 5)){
            return
        }
    }
    else{
        if(not ClickPic("Images/NoUsarUsoInterno.png", 425, 5)){
            return
        }
        WinWait, ACTUALIZACION DE ARTICULOS
        if(not ClickPic("Images/NoUsarUsoInterno_Seleccionado.png", 425, 5)){
            return
        }
    }
    WinWait, ACTUALIZACION DE ARTICULOS
    Send, %multiplied%
}

ActualizarDescripFecha(doAfter:="", replacement:=""){
    if(not WinExist("ACTUALIZACION DE ARTICULOS")){
        MsgBox No existe ACTUALIZACION DE ARTICULOS.
        return
    }
    
    if(replacement == ""){
        FormatTime, replacement, , MM/yyyy
    }
    
    if(WinExist("ahk_class TFrmBuscar")){ ;if(PicExists("Images/VentanaBuscar.png")){
        Send, {Enter}
    }
    WinActivate, ACTUALIZACION DE ARTICULOS
    
    WinWait, ACTUALIZACION DE ARTICULOS
    WinMenuSelectItem, ACTUALIZACION DE ARTICULOS, , Modificar
    WinWait, ACTUALIZACION DE ARTICULOS
    ;if(not ClickPic("Images/DescripAdicional.png", 180, 5)){return}
    ControlFocus, TEdit8, ACTUALIZACION DE ARTICULOS ;TEdit8 es la ID del campo de texto de Descripción Adicional.
    WinWait, ACTUALIZACION DE ARTICULOS
    SendMessage, 0x301, , , TEdit8, ACTUALIZACION DE ARTICULOS ;SendMessage, 0x301 envía CTRL+C. Por si accidentalmente sobreescribimos la descripción de un artículo equivocado.
    WinWait, ACTUALIZACION DE ARTICULOS
    Send, %replacement%
    WinWait, ACTUALIZACION DE ARTICULOS
    Send, {F10}
    Sleep, 150
    Send, {F10}
    Sleep, 150
    Send, {F10}
    
    if(doAfter == "search"){
        Sleep, 150
        Send, ^b ;Ctrl+B: Buscar
    }
    else if(doAfter == "next"){
        Sleep, 150
        ProximoArticulo()
    }
}

BuscarPorPortapapel(){
    if WinExist("OpenOffice Calc"){
        WinActivate, OpenOffice Calc
        WinWait, OpenOffice Calc
        if WinExist("Find & Replace"){
            WinActivate, Find & Replace
            WinWait, Find & Replace
            Send, !s ;Alt+S: Search For
        }
        else{
            Send, ^f ;Ctrl+F: Buscar
        }
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
    }
    if WinExist("Adobe Reader"){
        WinActivate, Adobe Reader
        WinWait, Adobe Reader
        Send, ^f ;Ctrl+F: Buscar
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
    }
}

SincronizarArticulosPrecio(){
    if(not WinExist("ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO")){
        MsgBox No existe ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO.
        return
    }
    if(not WinExist("ACTUALIZACION DE ARTICULOS")){
        MsgBox No existe ACTUALIZACION DE ARTICULOS.
        return
    }
    WinActivate, ACTUALIZACION DE ARTICULOS
    
    WinWait, ACTUALIZACION DE ARTICULOS
    WinMenuSelectItem, ACTUALIZACION DE ARTICULOS, , Copiar
    WinWait, ACTUALIZACION DE ARTICULOS
    tempClipboard := Clipboard ;Para "preservar" el portapapeles, usamos una variable auxiliar.
    Send, ^c{Esc} ;Ctrl+C+Enter: Copiar al portapapeles y Salir
    WinWait, ACTUALIZACION DE ARTICULOS
    tempArticleCode := Clipboard
    Clipboard = tempClipboard
    
    WinActivate, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    Send, ^b ;Ctrl+B: Buscar
    WinWait, ahk_class TFrmBuscar ;Esta es la ventana Buscar.
    Send, %tempArticleCode%
}

ProximoArticulo(){ 
    if(WinExist("ACTUALIZACION DE ARTICULOS")){
        ControlSend,,{PGDN}, ACTUALIZACION DE ARTICULOS
    }
    if(WinExist("ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO")){
        ControlSend,,{PGDN}, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    }
}

Media_Prev::
ActualizarDescripFecha("search")
return

^Media_Prev::
ActualizarDescripFecha("next")
return

Media_Next::
ProximoArticulo()
return

Launch_Mail::
SincronizarArticulosPrecio()
return

Browser_Search::
CopiarUnidadMedidaVentas()
return

^Browser_Search::
CopiarUnidadMedidaVentas()
Sleep,100
BuscarPorPortapapel()
return

Browser_Home::
WinActivate, ACTUALIZACION DE ARTICULOS ;HITLERS
PegarPrecio98o99(1.1495)
return

^Browser_Home::
WinActivate, ACTUALIZACION DE ARTICULOS ;HITLERS
PegarPrecio98o99(1.04975)
return

#IfWinActive SOS DE STOCK ; Works for EGRESOS and INGRESOS. AHK does not have an OR for this command.
::cdm::Cambio de mercadería - Blas
::cds::Corrección de stock - Blas
::mui::Uso Interno - Blas
#IfWinActive