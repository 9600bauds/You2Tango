#Include <ClickPic>

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force ; Close old versions of this script automatically.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

;{ Globals
global ventanaArticulos := "ACTUALIZACION DE ARTICULOS"
global campoMedidaVentas := "TEdit4"
global campoCodigoArt_Articulos := "TEdit11"
global campoDescAdicional := "TEdit8"
global campoDesc_Articulos := "TEdit9"

global ventanaPrecios := "ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO"
global campoCodigoArt_Precios_ModoNoModificar := "TEdit2" ;EN MODO NO MODIFICAR
global campoCodigoArt_Precios_ModoModificar := "TEdit6" ;EN MODO MODIFICAR
global campoPrecioActual := "TNumEditTg1"

global ventanaBuscar := "ahk_class TFrmBuscar"
global campoContenido_Buscar := "TcxCustomInnerTextEdit1"
global checkboxFiltrar := "TCheckBox4"
global checkboxIncremental := "TCheckBox3"

global ventanaNotepad := "ahk_class Notepad"

global multiplicadorPrecio1 := 1.21
global multiplicadorPrecio2 := 1
global multiplicadorExtra := 0
;}

;{ Ventana Artículos - Helpers
GetUnidadMedidaVentas(){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }

    ControlGetText, unidadMedida, %campoMedidaVentas%, %ventanaArticulos%
    return unidadMedida
}

CambiarCampoVentanaArticulos(field := "", newText = ""){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return false
    }
    
    if(WinExist(ventanaBuscar)){
        CerrarVentanaBuscar()
    }
    
    If(!IsAlwaysOnTop(ventanaArticulos)){
        WinActivate, %ventanaArticulos%
        WinWait, %ventanaArticulos%
    }
    
    WinMenuSelectItem, %ventanaArticulos%, , Modificar
    WinWait, %ventanaArticulos%
    ControlFocus, %field%, %ventanaArticulos% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    ControlSetText, %field%, %newText%, %ventanaArticulos%
    WinWait, %ventanaArticulos%
    Send, {F10}
    Sleep, 150
    Send, {F10}
    Sleep, 150
    Send, {F10}
    
    return true
}

GetCodigoVentanaArticulos(){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }
    
    ControlGetText, itemID, %campoCodigoArt_Articulos%, %ventanaArticulos%
    return itemID
}
;}

;{ Ventana Artículos - Funciones
EliminacionArticulo(doAfter:=""){    
    itemID := GetCodigoVentanaArticulos() ;Para el logging
    ControlGetText, oldDesc, %campoDesc_Articulos%, %ventanaArticulos% ;Para el logging.
    ControlGetText, Clipboard, %campoDesc_Articulos%, %ventanaArticulos% ;Copiamos al portapapeles, por si accidentalmente borramos un artículo equivocado.
    
    if(not CambiarCampoVentanaArticulos(campoDesc_Articulos, "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ")){
        return false
    }
    
    LogArticleDeletion(itemID, oldDesc)
       
    if(doAfter == "search"){
        Sleep, 150
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
    }
    else if(doAfter == "next"){
        Sleep, 150
        ProximoArticulo()
    }
}

ActualizarDescripFecha(doAfter:="", replacement:=""){   
    if(replacement == ""){
        FormatTime, replacement, , MM/yyyy
    }
        
    itemID := GetCodigoVentanaArticulos() ;Para el logging.
    ControlGetText, oldDesc, %campoDescAdicional%, %ventanaArticulos% ;Para el logging.
    ControlGetText, Clipboard, %campoDescAdicional%, %ventanaArticulos% ;Copiamos al portapapeles, por si accidentalmente sobreescribimos la descripción de un artículo equivocado.
    
    if(not CambiarCampoVentanaArticulos(campoDescAdicional, replacement)){
        return false
    }
    
    LogDescChange(itemID, oldDesc, replacement)
    
    if(doAfter == "search"){
        Sleep, 150
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
    }
    else if(doAfter == "next"){
        Sleep, 150
        ProximoArticulo()
    }
}

MassActualizarDesc(){
    arr = 1578,2134,3758,5638,6544,500,2886
    
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }
    
    if(not WinExist(ventanaBuscar)){
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
        WinWait, %ventanaBuscar%
    }
    
    Loop, parse, arr, `,,
    {
        ActualizarDescripFecha("search")
        Sleep, 250
        Send, %A_LoopField%
    }
    ActualizarDescripFecha()
}
;}

;{ Ventana Precios - Helpers
SeleccionarPrecio98o99(){
    if(not WinExist(ventanaPrecios)){
        MsgBox No existe %ventanaPrecios%.
        return false
    }
    
    If(!IsAlwaysOnTop(ventanaPrecios)){
        WinActivate, %ventanaPrecios%
        WinWait, %ventanaPrecios%
    }
    
    WinMenuSelectItem, %ventanaPrecios%, , Modificar
    WinWait, %ventanaPrecios%
    if(PicExists("Images/ActualizacionPrecios/Dolar.png")){
        if(not ClickPic("Images/ActualizacionPrecios/Dolar.png", 425, 5)){
            return false
        }
        WinWait, %ventanaPrecios%
        if(not ClickPic("Images/ActualizacionPrecios/Dolar_Seleccionado.png", 425, 5)){
            return false
        }
    }
    else{
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno.png", 425, 5)){
            return false
        }
        WinWait, %ventanaPrecios%
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno_Seleccionado.png", 425, 5)){
            return false
        }
    }
    WinWait, %ventanaPrecios%
    
    return true
}

GetCodigoVentanaPrecios(){
    if(not WinExist(ventanaPrecios)){
        MsgBox No existe %ventanaPrecios%.
        return
    }
    
    ControlGetText, itemID, %campoCodigoArt_Precios_ModoNoModificar%, %ventanaPrecios%
    if(itemID == "S" or itemID == "N"){
        ControlGetText, itemID, %campoCodigoArt_Precios_ModoModificar%, %ventanaPrecios%
    }
    return itemID
}
;}

;{ Ventana Precios - Funciones
IngresarMultiplicadoresPrecio(){
    explanation := "Ingrese el descuento con el siguiente formato:`n1.21 o 21% para +21`%`n0.8 o -20% para -20`%"
    newMultiplier := 0
    InputBox, newMultiplier, Descuento Básico, %explanation%,,,,,,,,%multiplicadorPrecio1%
    if ErrorLevel
        return ;Cancel
    if(ParsePercent(newMultiplier)) {
        multiplicadorPrecio1 := newMultiplier
    }
    else{
        MsgBox, No se ingresó un número. (%newMultiplier%)
    }
    
    InputBox, newMultiplier, Descuento Alternativo, %explanation%,,,,,,,,%multiplicadorPrecio2%
    if ErrorLevel
        return ;Cancel
    if(ParsePercent(newMultiplier)) {
        multiplicadorPrecio2 := newMultiplier
    }
    else{
        MsgBox, No se ingresó un número. (%newMultiplier%)
    }
}

PegarPrecio98o99(mult:=1){
    if(not SincronizadosArticulosPrecio()){
        MsgBox, Ventana Artículos y ventana Precios no están actualizadas!
        return false
    }
    
    if(not SeleccionarPrecio98o99()){
        return false
    }
    mult := ParsePercent(mult)
    if(mult == 0){
        MsgBox, Multiplicador inválido! - PegarPrecio98o99
        return false
    }
    
    Clipboard := RegExReplace(Clipboard, ",", ".") ;Reemplazar comas por puntos.
    Clipboard := RegExReplace(Clipboard, "\.(?![^.]+$)")  ;Quitar todos los puntos excepto el último.
    Clipboard := RegExReplace(Clipboard, "[^0-9.]") ;Eliminar todo excepto números y puntos.
    if(not IsNum(Clipboard)) {
        MsgBox, Clipboard is not a number.
        return
    }
    multiplied := (Clipboard * mult)
    multiplied = % Round(multiplied, 2) ;Tango sólo quiere 2 decimales.
    
    itemID := GetCodigoVentanaPrecios()
    ControlGetText, oldPrice, %campoPrecioActual%, %ventanaPrecios%
    
    percent := (100*multiplied/oldPrice)-100
    percent := Round(percent, 1)
    if(percent < -15 or percent > 20){
        MsgBox, 305, , Diferencia de %percent%`%, continuar? ;1+48+256
        IfMsgBox, Cancel
        {
            Send, {Esc}
            Sleep, 150
            Send, {F10}
            return
        }
    }
    LogPriceChange(itemID, oldPrice, multiplied)
    
    Send, %multiplied%
    Send, {F10}
    Sleep, 150
    Send, {F10}
}

MultiplicarPrecio98o99(mult:=0){
    if(mult == 0){
        explanation := "Ingrese el porcentaje a añadir o restar con el siguiente formato:`n1.21 o 21% para +21`%`n0.8 o -20% para -20`%"
        InputBox, multInput, Porcentaje, %explanation%,,,,,,,,%multiplicadorExtra%
        if ErrorLevel
            return ;Cancel
        mult := ParsePercent(multInput)
        if(!IsNum(mult) or mult == 0) {
            MsgBox, No se ingresó un número. (%mult%)
            return
        }
        multiplicadorExtra := multInput
    }
    
    if(not SeleccionarPrecio98o99()){
        return false
    }
    
    itemID := GetCodigoVentanaPrecios()
    ControlGetText, oldPrice, %campoPrecioActual%, %ventanaPrecios%
    multiplied := oldPrice*mult
    multiplied = % Round(multiplied, 2) ;Tango sólo quiere 2 decimales.
    
    LogPriceChange(itemID, oldPrice, multiplied)
    ControlSetText, %campoPrecioActual%, %multiplied%, %ventanaPrecios%
    Send, {F10}
    Sleep, 150
    Send, {F10}
}
;}

;{ Navegación
SincronizarArticulosPrecio(){
    if(not WinExist(ventanaPrecios)){
        MsgBox No existe %ventanaPrecios%.
        return
    }
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }
    if(WinExist(ventanaBuscar)){
        CerrarVentanaBuscar()
    }

    CodigoArticulo := GetCodigoVentanaArticulos()
    WinMenuSelectItem, %ventanaPrecios%, , Buscar, Por Clave
    WinWait, %ventanaBuscar% ;Ésta es la ventana Buscar.
    ControlSend, %campoContenido_Buscar%, %CodigoArticulo%, %ventanaBuscar% 
    
    CerrarVentanaBuscar()
}

SincronizadosArticulosPrecio(){
    if(not WinExist(ventanaPrecios)){
        MsgBox No existe %ventanaPrecios%.
        return false
    }
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return false
    }
    
    CodigoArticulos := GetCodigoVentanaArticulos()
    CodigoPrecios := GetCodigoVentanaPrecios()
    
    if(CodigoArticulos == CodigoPrecios){
        return true
    }
    else{
        return false
    }
}

ProximoArticulo(){ 
    if(WinExist(ventanaArticulos)){
        ;WinWait, %ventanaArticulos%
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Siguiente ;Modo Modificar.
        ;ControlSend,,{PGDN}, %ventanaArticulos%
    }
    if(WinExist(ventanaPrecios)){
        ;WinWait, %ventanaPrecios%
        WinMenuSelectItem, %ventanaPrecios%, , Buscar, Siguiente
        ;ControlSend,,{PGDN}, %ventanaPrecios%
    }
}

AnteriorArticulo(){
    if(WinExist(ventanaArticulos)){
        ControlSend,,{PGUP}, %ventanaArticulos%
    }
    if(WinExist(ventanaPrecios)){
        ControlSend,,{PGUP}, %ventanaPrecios%
    }
}
;}

;{ Ventana Buscar
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
        Sleep, 100
        if(PicExists("Images/OpenOfficeCalc/EndOf.png")){ ;Damn you, OpenOffice.
            Send, {Enter}
        }
    }
    if WinExist("Adobe Reader"){
        WinActivate, Adobe Reader
        WinWait, Adobe Reader
        Send, ^f ;Ctrl+F: Buscar
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
    }
}

CerrarVentanaBuscar(){
    if(not WinExist(ventanaBuscar)){
        MsgBox No existe %ventanaBuscar%.
        return
    }
    WinActivate, %ventanaBuscar%
    
    Control, Check, , %checkboxFiltrar%, %ventanaBuscar%
    Control, Uncheck, , %checkboxIncremental%, %ventanaBuscar%
    
    ControlGetText, CodigoIngresado, %campoContenido_Buscar%, %ventanaBuscar%
    if(CodigoIngresado == ""){
        Send, {Esc}
        WinWaitClose, %ventanaBuscar%
        return
    }
    
    Send, {Enter}
    If(WinExist(ventanaBuscar)){ ;Puede que ya hayamos apretado Enter nosotros.
        Send, {Enter}
    }
    WinWaitClose, %ventanaBuscar%
}
;}

;{ Logging
LogPriceChange(itemID := "", oldPrice := "", newPrice = ""){
    percent := (100*newPrice/oldPrice)-100
    percent := Round(percent, 1)
    finalText = %itemID%: %oldPrice% -> %newPrice% (%percent%`%)`r`n
    LogSend(finalText)
}

LogArticleDeletion(itemID := "", oldDesc := ""){
    finalText = Eliminación: %itemID% (%oldDesc%)`r`n
    LogSend(finalText)
}

LogDescChange(itemID := "", oldDesc := "", newDesc = ""){
    finalText = %itemID%: %oldDesc% -> %newDesc%`r`n
    LogSend(finalText)
}

LogSend(finalText := ""){
    if(not WinExist(ventanaNotepad)){
        prev := WinActive("A")
        Run, Notepad
        WinWait, %ventanaNotepad%
        WinActivate, ahk_id %prev%
    }
    ControlSend,,^{End}, %ventanaNotepad% ;Ctrl+End: Go to end of document
    Control, EditPaste, %finalText%, , %ventanaNotepad%
}
;}

;{ Misc
EstilizarVentanas(Activar := 1){
    if(Activar == 1){
        WinSet, AlwaysOnTop, On, %ventanaArticulos%
        WinSet, Region, 0-0 W572 H222, %ventanaArticulos% ;Máscara de 572x222 empezando en 0,0
        WinMove, %ventanaArticulos%, , 1028, 26, 572
        WinGetPos, X, Y, W, H, %ventanaArticulos%
        
        WinSet, AlwaysOnTop, On, %ventanaPrecios%

        WinMove, %ventanaPrecios%, , X, Y+222, W, H
        WinGetPos, X, Y, W, H, %ventanaPrecios%
        
        if(WinExist(ventanaNotepad)){
            WinSet, Region, 0-0 W572 H398, %ventanaPrecios% ;Máscara de 572x425 empezando en 0,0
            
            WinSet, Region, 0-0 W999 H999, %ventanaNotepad% ;Literalmente sólo para que tenga los 3 pixeles negros feos
            WinSet, AlwaysOnTop, On, %ventanaNotepad%
            WinMove, %ventanaNotepad%, , X, Y+398 , W, 125
        }
        

    }
    else{
        WinSet, AlwaysOnTop, Off, %ventanaArticulos%
        WinSet, Region, , %ventanaArticulos%
        
        WinSet, AlwaysOnTop, Off, %ventanaPrecios%
        WinSet, Region, , %ventanaPrecios%
        
        WinSet, AlwaysOnTop, Off, %ventanaNotepad%
        WinSet, Region, , %ventanaNotepad%
    }
}

ParsePercent(input){
    if(InStr(input, "%")){
        input := RegExReplace(input, "[^0-9|\-|.]") ;Sólo numeros.
        return (100+input)/100
    }
    else{
        if(not IsNum(input)){
            return 0
        }
        return input
    }
}

IsNum( str ) { ;Fuck AHK.
	if str is number
		return true
	return false
}

IsAlwaysOnTop( Window ) {
    WinGet, Estilo, ExStyle, %Window%
    Return (Estilo & 0x8) ; 0x8 is WS_EX_TOPMOST.
}
;}

;{ AUTOEXEC
if(not WinExist(ventanaNotepad)){
    Run, Notepad
}
;}

Launch_Media::
;EliminacionArticulo()
MsgBox, Testing...
return

Volume_Up::
EstilizarVentanas(1)
return

Volume_Down::
EstilizarVentanas(0)
return

Volume_Mute::
IngresarMultiplicadoresPrecio()
return

^Volume_Mute::
MultiplicarPrecio98o99()
return

Media_Play_Pause::
ActualizarDescripFecha("search")
return

^Media_Play_Pause::
ActualizarDescripFecha("next")
return

Media_Prev::
AnteriorArticulo()
return

Media_Next::
ProximoArticulo()
return

Launch_Mail::
SincronizarArticulosPrecio()
return

^Launch_Mail::
SincronizarArticulosPrecio()
Sleep, 100
Clipboard := GetUnidadMedidaVentas()
Sleep, 100
BuscarPorPortapapel()
return

Browser_Search::
Clipboard := GetUnidadMedidaVentas()
Sleep,100
BuscarPorPortapapel()

;Sleep,100
;Send, {Esc}
;Send, {Left}
;Send, {Left}
;Send, ^c

return

Browser_Home::
PegarPrecio98o99(multiplicadorPrecio1)
return

^Browser_Home::
PegarPrecio98o99(multiplicadorPrecio2)
return

#IfWinActive SOS DE STOCK ; Works for EGRESOS and INGRESOS. AHK does not have an OR operand for this command.
::cdm::Cambio de Mercadería - Blas
::cds::Corrección de Stock - Blas
::mui::Roto/Uso Interno - Blas
::-b:: - Blas
#IfWinActive