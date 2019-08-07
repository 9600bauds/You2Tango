#Include <ClickPic>

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force ; Close old versions of this script automatically.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

global ventanaArticulos := "ACTUALIZACION DE ARTICULOS"
global campoMedidaVentas := "TEdit4"
global campoCodigoArt_Articulos := "TEdit11"
global campoDescAdicional := "TEdit8"
global campoDesc_Articulos := "TEdit9"

global ventanaPrecios := "ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO"
global campoCodigoArt_Precios := "TEdit6"
global campoPrecioActual := "TNumEditTg1"

global ventanaBuscar := "ahk_class TFrmBuscar"
global campoContenido_Buscar := "TcxCustomInnerTextEdit1"
global checkboxFiltrar := "TCheckBox4"
global checkboxIncremental := "TCheckBox3"

global ventanaNotepad := "ahk_class Notepad"

CopiarUnidadMedidaVentas(){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }

    ControlGetText, Clipboard, %campoMedidaVentas%, %ventanaArticulos%
}

PegarPrecio98o99(mult:=1){
    Clipboard := RegExReplace(Clipboard, ",", ".") ;Reemplazar comas por puntos.
    Clipboard := RegExReplace(Clipboard, "\.(?![^.]+$)")  ;Quitar todos los puntos excepto el último.
    Clipboard := RegExReplace(Clipboard, "[^0-9.]") ;Eliminar todo excepto números y puntos.
    if(not IsNum(Clipboard)) {
        MsgBox, Clipboard is not a number.
        return
    }
    multiplied := (Clipboard * mult)
    multiplied = % Round(multiplied, 2) ;Tango sólo quiere 2 decimales.
   
    if(not WinExist(ventanaPrecios)){
        MsgBox No existe %ventanaPrecios%.
        return
    }
    
    If(!IsAlwaysOnTop(ventanaPrecios)){
        WinActivate, %ventanaPrecios%
        WinWait, %ventanaPrecios%
    }
    
    WinMenuSelectItem, %ventanaPrecios%, , Modificar
    WinWait, %ventanaPrecios%
    if(PicExists("Images/ActualizacionPrecios/Dolar.png")){
        if(not ClickPic("Images/ActualizacionPrecios/Dolar.png", 425, 5)){
            return
        }
        WinWait, %ventanaPrecios%
        if(not ClickPic("Images/ActualizacionPrecios/Dolar_Seleccionado.png", 425, 5)){
            return
        }
    }
    else{
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno.png", 425, 5)){
            return
        }
        WinWait, %ventanaPrecios%
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno_Seleccionado.png", 425, 5)){
            return
        }
    }
    WinWait, %ventanaPrecios%
    
    ControlGetText, itemID, %campoCodigoArt_Precios%, %ventanaPrecios%
    ControlGetText, oldPrice, %campoPrecioActual%, %ventanaPrecios%
    
    percent := (100*multiplied/oldPrice)-100
    percent := Round(percent, 1)
    if(percent < -15 or percent > 20)
        MsgBox, 305, , Diferencia de %percent%`%, continuar? ;1+48+256
        IfMsgBox, Cancel
        {
            Send, {Esc}
            Sleep, 150
            Send, {F10}
            return
        }
    LogPriceChange(itemID, oldPrice, multiplied)
    
    Send, %multiplied%
    Send, {F10}
    Sleep, 150
    Send, {F10}
}

ActualizarDescripFecha(doAfter:="", replacement:=""){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }
    
    if(replacement == ""){
        FormatTime, replacement, , MM/yyyy
    }
    
    if(WinExist(ventanaBuscar)){
        CerrarVentanaBuscar()
    }
    
    If(!IsAlwaysOnTop(ventanaArticulos)){
        WinActivate, %ventanaArticulos%
        WinWait, %ventanaArticulos%
    }
    
    ControlGetText, Clipboard, %campoDescAdicional%, %ventanaArticulos% ;Copiamos al portapapeles, por si accidentalmente sobreescribimos la descripción de un artículo equivocado.
    WinMenuSelectItem, %ventanaArticulos%, , Modificar
    WinWait, %ventanaArticulos%
    ControlFocus, %campoDescAdicional%, %ventanaArticulos% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    ControlSetText, %campoDescAdicional%, %replacement%, %ventanaArticulos%
    WinWait, %ventanaArticulos%
    Send, {F10}
    Sleep, 150
    Send, {F10}
    Sleep, 150
    Send, {F10}
    
    if(doAfter == "search"){
        Sleep, 150
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
    }
    else if(doAfter == "next"){
        Sleep, 150
        ProximoArticulo()
    }
}

EliminacionArticulo(doAfter:=""){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }
    
    if(WinExist(ventanaBuscar)){
        CerrarVentanaBuscar()
    }
    
    If(!IsAlwaysOnTop(ventanaArticulos)){
        WinActivate, %ventanaArticulos%
        WinWait, %ventanaArticulos%
    }
    
    ControlGetText, Clipboard, %campoDesc_Articulos%, %ventanaArticulos% ;Copiamos al portapapeles, por si accidentalmente borramos un artículo equivocado.
    
    ControlGetText, itemID, %campoCodigoArt_Articulos%, %ventanaArticulos%
    LogArticleDeletion(itemID)

    WinMenuSelectItem, %ventanaArticulos%, , Modificar
    WinWait, %ventanaArticulos%
    ControlFocus, %campoDesc_Articulos%, %ventanaArticulos% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    ControlSetText, %campoDesc_Articulos%, ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ, %ventanaArticulos%
    WinWait, %ventanaArticulos%
    Send, {F10}
    Sleep, 150
    Send, {F10}
    Sleep, 150
    Send, {F10}
    
    if(doAfter == "search"){
        Sleep, 150
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
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

    ControlGetText, CodigoArticulo, %campoCodigoArt_Articulos%, %ventanaArticulos%
    WinMenuSelectItem, %ventanaPrecios%, , Buscar, Por Clave
    WinWait, %ventanaBuscar% ;Ésta es la ventana Buscar.
    ControlSend, %campoContenido_Buscar%, %CodigoArticulo%, %ventanaBuscar% 
    
    CerrarVentanaBuscar()
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

LogPriceChange(itemID := "", oldPrice := "", newPrice = ""){
    percent := (100*newPrice/oldPrice)-100
    percent := Round(percent, 1)
    finalText = %itemID%: %oldPrice% -> %newPrice% (%percent%`%)`r`n
    Control, EditPaste, %finalText%, , %ventanaNotepad%
}

LogArticleDeletion(itemID := ""){
    finalText = Eliminación: %itemID%`r`n
    Control, EditPaste, %finalText%, , %ventanaNotepad%
}

Launch_Media::
MsgBox, Testing...
return

Volume_Up::
EstilizarVentanas(1)
return

Volume_Down::
EstilizarVentanas(0)
return

;Volume_Mute::

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
CopiarUnidadMedidaVentas()
Sleep, 100
BuscarPorPortapapel()
return

Browser_Search::
CopiarUnidadMedidaVentas()
Sleep,100
BuscarPorPortapapel()
return

Browser_Home::
PegarPrecio98o99(1.21)
return

^Browser_Home::
PegarPrecio98o99(1.21/45)
return

#IfWinActive SOS DE STOCK ; Works for EGRESOS and INGRESOS. AHK does not have an OR operand for this command.
::cdm::Cambio de Mercadería - Blas
::cds::Corrección de Stock - Blas
::mui::Roto/Uso Interno - Blas
::-b:: - Blas
#IfWinActive